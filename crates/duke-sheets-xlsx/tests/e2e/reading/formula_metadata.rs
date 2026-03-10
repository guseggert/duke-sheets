use std::fs::File;
use std::io::Write;

use crate::{cleanup_fixture, temp_fixture_path};
use duke_sheets_core::{CellAddress, CellError};
use duke_sheets_xlsx::XlsxReader;

fn write_single_sheet_fixture(path: &std::path::Path, sheet_xml: &str) {
    let file = File::create(path).expect("create fixture file");
    let mut zip = zip::ZipWriter::new(file);
    let options = zip::write::SimpleFileOptions::default();

    zip.start_file("[Content_Types].xml", options)
        .expect("content types part");
    zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#)
        .expect("write content types");

    zip.start_file("_rels/.rels", options)
        .expect("root rels part");
    zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#)
        .expect("write root rels");

    zip.start_file("xl/workbook.xml", options)
        .expect("workbook part");
    zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#)
        .expect("write workbook");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("workbook rels part");
    zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#)
        .expect("write workbook rels");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("sheet part");
    zip.write_all(sheet_xml.as_bytes())
        .expect("write sheet xml");

    zip.finish().expect("finish zip");
}

fn formula_text_at<'a>(sheet: &'a duke_sheets_core::Worksheet, address: &str) -> Option<&'a str> {
    let addr = CellAddress::parse(address).expect("valid address");
    sheet.get_formula_at(addr.row, addr.col)
}

#[test]
fn test_shared_formula_follower_materialized() {
    let path = temp_fixture_path();
    let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="n"><v>1</v></c>
      <c r="B1" t="n"><v>2</v></c>
      <c r="D1"><f t="shared" si="3">SUM($A$1:B1)+LEN("A1")</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="n"><v>4</v></c>
      <c r="B2" t="n"><v>5</v></c>
      <c r="D2"><f t="shared" si="3"/><v>9</v></c>
    </row>
  </sheetData>
</worksheet>"#;
    write_single_sheet_fixture(&path, sheet_xml);

    let workbook = XlsxReader::read_file(&path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("sheet exists");

    assert_eq!(
        formula_text_at(sheet, "D1"),
        Some("=SUM($A$1:B1)+LEN(\"A1\")")
    );
    assert_eq!(
        formula_text_at(sheet, "D2"),
        Some("=SUM($A$1:B2)+LEN(\"A1\")")
    );

    cleanup_fixture(&path);
}

#[test]
fn test_datatable_formula_placeholder_and_cached_value() {
    let path = temp_fixture_path();
    let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="dataTable" ref="A1:B2" r1="C1" r2="C2"/><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;
    write_single_sheet_fixture(&path, sheet_xml);

    let workbook = XlsxReader::read_file(&path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("sheet exists");

    assert_eq!(formula_text_at(sheet, "A1"), Some("=TABLE(C1,C2)"));
    assert_eq!(sheet.get_value("A1").unwrap().as_number(), Some(42.0));

    cleanup_fixture(&path);
}

#[test]
fn test_outline_and_sheet_view_metadata() {
    let path = temp_fixture_path();
    let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0" tabSelected="1">
      <pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r="2" outlineLevel="2" collapsed="1"><c r="A2" t="n"><v>1</v></c></row>
  </sheetData>
  <cols>
    <col min="3" max="3" outlineLevel="3" collapsed="1"/>
  </cols>
</worksheet>"#;
    write_single_sheet_fixture(&path, sheet_xml);

    let workbook = XlsxReader::read_file(&path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("sheet exists");

    assert!(sheet.is_selected());
    assert_eq!(
        sheet.freeze_panes().map(|fp| (fp.row, fp.col)),
        Some((3, 2))
    );
    assert_eq!(sheet.row_outline_level(1), 2);
    assert!(sheet.is_row_collapsed(1));
    assert_eq!(sheet.column_outline_level(2), 3);
    assert!(sheet.is_column_collapsed(2));

    cleanup_fixture(&path);
}

/// Test that the reader correctly parses cm attributes on cells.
/// This simulates an Excel-generated file with dynamic array metadata.
#[test]
fn test_reader_parses_cm_attribute() {
    let path = temp_fixture_path();
    // Handcrafted sheet XML with cm attributes (like Excel would produce)
    let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1"><f>SEQUENCE(3,1)</f><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" cm="2"><v>2</v></c>
    </row>
    <row r="3">
      <c r="A3" cm="2"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>"#;
    write_single_sheet_fixture(&path, sheet_xml);

    let workbook = XlsxReader::read_file(&path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("sheet exists");

    // A1 is a formula (anchor cell with cm=1)
    assert_eq!(formula_text_at(sheet, "A1"), Some("=SEQUENCE(3,1)"));
    assert_eq!(sheet.get_value("A1").unwrap().as_number(), Some(1.0));

    // Ghost cells are now SpillTarget (reader reconstructs dynamic array)
    assert!(sheet.get_value("A2").unwrap().is_spill_target());
    assert!(sheet.get_value("A3").unwrap().is_spill_target());
    // Resolved values match
    assert_eq!(sheet.get_value_at(1, 0).as_number(), Some(2.0));
    assert_eq!(sheet.get_value_at(2, 0).as_number(), Some(3.0));

    cleanup_fixture(&path);
}

/// Test reader with cm attribute and string-type ghost cells.
#[test]
fn test_reader_parses_cm_attribute_string_ghost() {
    let path = temp_fixture_path();
    let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1" t="str"><f>UNIQUE(B1:B3)</f><v>apple</v></c>
    </row>
    <row r="2">
      <c r="A2" cm="2" t="str"><v>banana</v></c>
    </row>
    <row r="3">
      <c r="A3" cm="2" t="str"><v>cherry</v></c>
    </row>
  </sheetData>
</worksheet>"#;
    write_single_sheet_fixture(&path, sheet_xml);

    let workbook = XlsxReader::read_file(&path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("sheet exists");

    // A1 is formula with string cached value
    assert_eq!(formula_text_at(sheet, "A1"), Some("=UNIQUE(B1:B3)"));

    // Ghost cells are SpillTarget; resolved values are strings
    assert!(sheet.get_value("A2").unwrap().is_spill_target());
    assert!(sheet.get_value("A3").unwrap().is_spill_target());
    assert_eq!(sheet.get_value_at(1, 0).as_string(), Some("banana"));
    assert_eq!(sheet.get_value_at(2, 0).as_string(), Some("cherry"));

    cleanup_fixture(&path);
}

/// Test reader with cm attribute on error-type cells.
#[test]
fn test_reader_parses_cm_attribute_error_anchor() {
    let path = temp_fixture_path();
    let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1" t="e"><f>SEQUENCE(3)</f><v>#SPILL!</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>999</v></c>
    </row>
  </sheetData>
</worksheet>"#;
    write_single_sheet_fixture(&path, sheet_xml);

    let workbook = XlsxReader::read_file(&path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("sheet exists");

    // A1 is formula with #SPILL! cached error
    let a1 = sheet.get_value("A1").unwrap();
    assert!(formula_text_at(sheet, "A1").is_some());
    assert!(matches!(
        a1,
        duke_sheets_core::CellValue::Error(CellError::Spill)
    ));

    // A2 is the blocker value
    assert_eq!(sheet.get_value("A2").unwrap().as_number(), Some(999.0));

    cleanup_fixture(&path);
}
