use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use duke_sheets_core::table::{Table, TableColumn, TableStyleInfo, TotalsRowFunction};
use duke_sheets_core::CellRange;

/// Read a table definition from `xl/tables/tableN.xml`.
pub(crate) fn read_table<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    table_path: &str,
) -> XlsxResult<Option<Table>> {
    let file = match archive.by_name(table_path) {
        Ok(f) => f,
        Err(_) => return Ok(None),
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut table: Option<Table> = None;
    let mut columns: Vec<TableColumn> = Vec::new();
    let mut style_info: Option<TableStyleInfo> = None;

    // State for parsing child elements of tableColumn
    let mut current_column: Option<TableColumn> = None;
    let mut in_calculated_column_formula = false;
    let mut in_totals_row_formula = false;
    let mut formula_text = String::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"table" if table.is_none() => table = Some(parse_table_attrs(&e)?),
                b"tableColumn" => current_column = Some(parse_table_column_attrs(&e)?),
                b"calculatedColumnFormula" => {
                    in_calculated_column_formula = true;
                    formula_text.clear();
                }
                b"totalsRowFormula" => {
                    in_totals_row_formula = true;
                    formula_text.clear();
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"table" if table.is_none() => table = Some(parse_table_attrs(&e)?),
                b"tableColumn" => {
                    let col = parse_table_column_attrs(&e)?;
                    columns.push(col);
                }
                b"tableStyleInfo" => style_info = Some(parse_table_style_info(&e)),
                b"autoFilter" => {
                    // We note the autoFilter exists but don't need to parse
                    // filter criteria for now - the ref is on the table itself.
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_calculated_column_formula || in_totals_row_formula {
                    formula_text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"tableColumn" => {
                    if let Some(col) = current_column.take() {
                        columns.push(col);
                    }
                }
                b"calculatedColumnFormula" => {
                    if let Some(ref mut col) = current_column {
                        col.calculated_column_formula = Some(formula_text.clone());
                    }
                    in_calculated_column_formula = false;
                }
                b"totalsRowFormula" => {
                    if let Some(ref mut col) = current_column {
                        col.totals_row_formula = Some(formula_text.clone());
                    }
                    in_totals_row_formula = false;
                }
                b"table" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    if let Some(ref mut t) = table {
        t.columns = columns;
        t.style_info = style_info;
    }

    Ok(table)
}

/// Parse the `<table>` element attributes.
fn parse_table_attrs(e: &quick_xml::events::BytesStart) -> XlsxResult<Table> {
    let mut id: u32 = 0;
    let mut name = String::new();
    let mut display_name = String::new();
    let mut reference = String::new();
    let mut header_row_count: u32 = 1; // default per spec
    let mut totals_row_count: u32 = 0;
    let mut totals_row_shown = true; // default per spec

    for attr in e.attributes().flatten() {
        let key = attr.key.local_name();
        let val = attr.unescape_value().unwrap_or_default();
        match key.as_ref() {
            b"id" => id = val.parse().unwrap_or(0),
            b"name" => name = val.into_owned(),
            b"displayName" => display_name = val.into_owned(),
            b"ref" => reference = val.into_owned(),
            b"headerRowCount" => header_row_count = val.parse().unwrap_or(1),
            b"totalsRowCount" => totals_row_count = val.parse().unwrap_or(0),
            b"totalsRowShown" => totals_row_shown = val.as_ref() != "0",
            _ => {}
        }
    }

    if display_name.is_empty() {
        display_name = name.clone();
    }

    let cell_range = CellRange::parse(&reference)
        .map_err(|_| XlsxError::InvalidFormat(format!("Bad table ref: {reference}")))?;

    Ok(Table {
        id,
        name,
        display_name,
        reference: cell_range,
        columns: Vec::new(),
        style_info: None,
        header_row_count,
        totals_row_count,
        totals_row_shown,
    })
}

/// Parse `<tableColumn>` attributes.
fn parse_table_column_attrs(e: &quick_xml::events::BytesStart) -> XlsxResult<TableColumn> {
    let mut id: u32 = 0;
    let mut name = String::new();
    let mut totals_row_function: Option<TotalsRowFunction> = None;
    let mut totals_row_label: Option<String> = None;

    for attr in e.attributes().flatten() {
        let key = attr.key.local_name();
        let val = attr.unescape_value().unwrap_or_default();
        match key.as_ref() {
            b"id" => id = val.parse().unwrap_or(0),
            b"name" => name = val.into_owned(),
            b"totalsRowFunction" => {
                totals_row_function = TotalsRowFunction::from_ooxml(&val);
            }
            b"totalsRowLabel" => totals_row_label = Some(val.into_owned()),
            _ => {}
        }
    }

    Ok(TableColumn {
        id,
        name,
        totals_row_function,
        totals_row_formula: None,
        totals_row_label,
        calculated_column_formula: None,
    })
}

/// Parse `<tableStyleInfo>` attributes.
fn parse_table_style_info(e: &quick_xml::events::BytesStart) -> TableStyleInfo {
    let mut info = TableStyleInfo::default();

    for attr in e.attributes().flatten() {
        let key = attr.key.local_name();
        let val = attr.unescape_value().unwrap_or_default();
        match key.as_ref() {
            b"name" => info.name = Some(val.into_owned()),
            b"showFirstColumn" => info.show_first_column = val.as_ref() == "1",
            b"showLastColumn" => info.show_last_column = val.as_ref() == "1",
            b"showRowStripes" => info.show_row_stripes = val.as_ref() == "1",
            b"showColumnStripes" => info.show_column_stripes = val.as_ref() == "1",
            _ => {}
        }
    }

    info
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Helper: create a minimal zip archive with one XML entry.
    fn zip_with_entry(path: &str, xml: &str) -> zip::ZipArchive<std::io::Cursor<Vec<u8>>> {
        use std::io::{Cursor, Write};
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(path, options).unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        let cursor = zip.finish().unwrap();
        zip::ZipArchive::new(cursor).unwrap()
    }

    #[test]
    fn test_read_basic_table() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="Sales" displayName="Sales" ref="A1:C5"
       headerRowCount="1" totalsRowCount="0">
  <autoFilter ref="A1:C5"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Product"/>
    <tableColumn id="2" name="Region"/>
    <tableColumn id="3" name="Revenue"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2"
      showFirstColumn="0" showLastColumn="0"
      showRowStripes="1" showColumnStripes="0"/>
</table>"#;

        let mut archive = zip_with_entry("xl/tables/table1.xml", xml);
        let table = read_table(&mut archive, "xl/tables/table1.xml")
            .unwrap()
            .unwrap();

        assert_eq!(table.id, 1);
        assert_eq!(table.name, "Sales");
        assert_eq!(table.display_name, "Sales");
        assert_eq!(table.reference.to_string(), "A1:C5");
        assert_eq!(table.header_row_count, 1);
        assert_eq!(table.totals_row_count, 0);
        assert!(table.has_header_row());
        assert!(!table.has_totals_row());

        assert_eq!(table.columns.len(), 3);
        assert_eq!(table.columns[0].name, "Product");
        assert_eq!(table.columns[1].name, "Region");
        assert_eq!(table.columns[2].name, "Revenue");

        let style = table.style_info.unwrap();
        assert_eq!(style.name.as_deref(), Some("TableStyleMedium2"));
        assert!(!style.show_first_column);
        assert!(style.show_row_stripes);
    }

    #[test]
    fn test_read_table_with_totals_and_formulas() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="2" name="Data" displayName="Data" ref="B2:D8"
       totalsRowCount="1">
  <autoFilter ref="B2:D7"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Item" totalsRowLabel="Total"/>
    <tableColumn id="2" name="Qty" totalsRowFunction="sum"/>
    <tableColumn id="3" name="Price" totalsRowFunction="custom">
      <totalsRowFormula>SUBTOTAL(109,[Price])</totalsRowFormula>
    </tableColumn>
  </tableColumns>
</table>"#;

        let mut archive = zip_with_entry("xl/tables/table2.xml", xml);
        let table = read_table(&mut archive, "xl/tables/table2.xml")
            .unwrap()
            .unwrap();

        assert_eq!(table.id, 2);
        assert_eq!(table.name, "Data");
        assert!(table.has_totals_row());
        assert_eq!(table.totals_row_count, 1);

        assert_eq!(table.columns[0].totals_row_label.as_deref(), Some("Total"));
        assert_eq!(
            table.columns[1].totals_row_function,
            Some(TotalsRowFunction::Sum)
        );
        assert_eq!(
            table.columns[2].totals_row_function,
            Some(TotalsRowFunction::Custom)
        );
        assert_eq!(
            table.columns[2].totals_row_formula.as_deref(),
            Some("SUBTOTAL(109,[Price])")
        );
    }

    #[test]
    fn test_read_table_with_calculated_column() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="3" name="Calc" displayName="Calc" ref="A1:C4">
  <tableColumns count="3">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
    <tableColumn id="3" name="Total">
      <calculatedColumnFormula>[A]+[B]</calculatedColumnFormula>
    </tableColumn>
  </tableColumns>
</table>"#;

        let mut archive = zip_with_entry("xl/tables/table3.xml", xml);
        let table = read_table(&mut archive, "xl/tables/table3.xml")
            .unwrap()
            .unwrap();

        assert_eq!(table.columns[2].name, "Total");
        assert_eq!(
            table.columns[2].calculated_column_formula.as_deref(),
            Some("[A]+[B]")
        );
        // headerRowCount defaults to 1
        assert_eq!(table.header_row_count, 1);
    }

    #[test]
    fn test_read_table_missing_file() {
        let xml = r#"<?xml version="1.0"?><dummy/>"#;
        let mut archive = zip_with_entry("other.xml", xml);
        let result = read_table(&mut archive, "xl/tables/table1.xml").unwrap();
        assert!(result.is_none());
    }
}
