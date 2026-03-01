use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use duke_sheets_core::CellAddress;

use super::{write_xml_part, XlsxResult, NS_SPREADSHEET};

pub(super) fn write_table_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    sheet: &duke_sheets_core::Worksheet,
    table_in_sheet_idx: usize,
    global_num: usize,
) -> XlsxResult<()> {
    let table = &sheet.tables()[table_in_sheet_idx];
    let path = format!("xl/tables/table{}.xml", global_num);

    write_xml_part(zip, &path, |w| {
        let mut tag = BytesStart::new("table");
        tag.push_attribute(("xmlns", NS_SPREADSHEET));

        let id_str = table.id.to_string();
        tag.push_attribute(("id", id_str.as_str()));
        tag.push_attribute(("name", table.name.as_str()));
        tag.push_attribute(("displayName", table.display_name.as_str()));

        let ref_str = table.reference.to_string();
        tag.push_attribute(("ref", ref_str.as_str()));

        if table.header_row_count != 1 {
            let hrc = table.header_row_count.to_string();
            tag.push_attribute(("headerRowCount", hrc.as_str()));
        }

        if table.totals_row_count > 0 {
            let trc = table.totals_row_count.to_string();
            tag.push_attribute(("totalsRowCount", trc.as_str()));
        }

        if !table.totals_row_shown {
            tag.push_attribute(("totalsRowShown", "0"));
        }

        w.write_event(Event::Start(tag))?;

        {
            let auto_ref = if table.has_totals_row() {
                let start = &table.reference.start;
                let end_row = table
                    .reference
                    .end
                    .row
                    .saturating_sub(table.totals_row_count);
                let end_col = table.reference.end.col;
                let end_addr = CellAddress::new(end_row, end_col);
                format!("{}:{}", start, end_addr)
            } else {
                ref_str.clone()
            };
            let mut af = BytesStart::new("autoFilter");
            af.push_attribute(("ref", auto_ref.as_str()));
            w.write_event(Event::Empty(af))?;
        }

        if !table.columns.is_empty() {
            let count_str = table.columns.len().to_string();
            let mut tc = BytesStart::new("tableColumns");
            tc.push_attribute(("count", count_str.as_str()));
            w.write_event(Event::Start(tc))?;

            for col in &table.columns {
                let col_id = col.id.to_string();
                let has_child =
                    col.calculated_column_formula.is_some() || col.totals_row_formula.is_some();

                let mut tc_el = BytesStart::new("tableColumn");
                tc_el.push_attribute(("id", col_id.as_str()));
                tc_el.push_attribute(("name", col.name.as_str()));

                if let Some(ref label) = col.totals_row_label {
                    tc_el.push_attribute(("totalsRowLabel", label.as_str()));
                }
                if let Some(func) = col.totals_row_function {
                    tc_el.push_attribute(("totalsRowFunction", func.to_ooxml()));
                }

                if has_child {
                    w.write_event(Event::Start(tc_el))?;
                    if let Some(ref formula) = col.totals_row_formula {
                        w.create_element("totalsRowFormula")
                            .write_text_content(BytesText::new(formula))?;
                    }
                    if let Some(ref formula) = col.calculated_column_formula {
                        w.create_element("calculatedColumnFormula")
                            .write_text_content(BytesText::new(formula))?;
                    }
                    w.write_event(Event::End(BytesEnd::new("tableColumn")))?;
                } else {
                    w.write_event(Event::Empty(tc_el))?;
                }
            }

            w.write_event(Event::End(BytesEnd::new("tableColumns")))?;
        }

        if let Some(ref style) = table.style_info {
            let mut si = BytesStart::new("tableStyleInfo");
            if let Some(ref name) = style.name {
                si.push_attribute(("name", name.as_str()));
            }
            si.push_attribute((
                "showFirstColumn",
                if style.show_first_column { "1" } else { "0" },
            ));
            si.push_attribute((
                "showLastColumn",
                if style.show_last_column { "1" } else { "0" },
            ));
            si.push_attribute((
                "showRowStripes",
                if style.show_row_stripes { "1" } else { "0" },
            ));
            si.push_attribute((
                "showColumnStripes",
                if style.show_column_stripes { "1" } else { "0" },
            ));
            w.write_event(Event::Empty(si))?;
        }

        w.write_event(Event::End(BytesEnd::new("table")))?;
        Ok(())
    })
}
