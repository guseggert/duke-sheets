use std::fs::File;
use std::io::{Cursor, Read, Write};

use crate::{cleanup_fixture, temp_fixture_path};
use duke_sheets_xlsx::{XlsxReader, XlsxWriter};
use quick_xml::events::Event;
use quick_xml::reader::Reader;

fn write_fixture(path: &std::path::Path) {
    let file = File::create(path).expect("create fixture file");
    let mut zip = zip::ZipWriter::new(file);
    let options = zip::write::SimpleFileOptions::default();

    zip.start_file("[Content_Types].xml", options)
        .expect("content types part");
    zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#)
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
    zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#)
        .expect("write workbook rels");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("sheet part");
    zip.write_all(br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1" t="n"><v>42</v></c></row></sheetData></worksheet>"#)
        .expect("write worksheet");

    zip.start_file("xl/styles.xml", options)
        .expect("styles part");
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><name val="Calibri"/><color rgb="FF008000"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="44" fontId="2" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1" applyFont="1"/>
  </cellXfs>
  <cellStyles count="3">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="1"/>
    <cellStyle name="Currency" xfId="2" builtinId="4"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#,
    )
    .expect("write styles");

    zip.finish().expect("finish zip");
}

fn read_styles_xml_from_xlsx(buf: &[u8]) -> String {
    let mut zip = zip::ZipArchive::new(Cursor::new(buf)).expect("open zip");
    let mut styles_file = zip.by_name("xl/styles.xml").expect("styles.xml part");
    let mut styles_xml = String::new();
    styles_file
        .read_to_string(&mut styles_xml)
        .expect("read styles.xml");
    styles_xml
}

#[derive(Default)]
struct StylesSnapshot {
    cell_style_xf_count: usize,
    named_styles: Vec<String>,
    cell_xf_ids: Vec<u32>,
}

fn parse_styles_snapshot(styles_xml: &str) -> StylesSnapshot {
    let mut reader = Reader::from_str(styles_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut snapshot = StylesSnapshot::default();
    let mut in_cell_style_xfs = false;
    let mut in_cell_xfs = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"cellStyleXfs" => in_cell_style_xfs = true,
                b"cellXfs" => in_cell_xfs = true,
                b"cellStyle" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"name" {
                            if let Ok(name) = attr.unescape_value() {
                                snapshot.named_styles.push(name.to_string());
                            }
                        }
                    }
                }
                b"xf" if in_cell_style_xfs => {
                    snapshot.cell_style_xf_count += 1;
                }
                b"xf" if in_cell_xfs => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"xfId" {
                            if let Ok(xf_id) = attr.unescape_value() {
                                if let Ok(xf_id) = xf_id.parse::<u32>() {
                                    snapshot.cell_xf_ids.push(xf_id);
                                }
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"cellStyle" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"name" {
                            if let Ok(name) = attr.unescape_value() {
                                snapshot.named_styles.push(name.to_string());
                            }
                        }
                    }
                }
                b"xf" if in_cell_style_xfs => {
                    snapshot.cell_style_xf_count += 1;
                }
                b"xf" if in_cell_xfs => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"xfId" {
                            if let Ok(xf_id) = attr.unescape_value() {
                                if let Ok(xf_id) = xf_id.parse::<u32>() {
                                    snapshot.cell_xf_ids.push(xf_id);
                                }
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"cellStyleXfs" => in_cell_style_xfs = false,
                b"cellXfs" => in_cell_xfs = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => panic!("styles parse error: {e}"),
            _ => {}
        }
        buf.clear();
    }

    snapshot
}

#[test]
fn test_roundtrip_preserves_cell_style_xfs_and_named_styles() {
    let path = temp_fixture_path();
    write_fixture(&path);

    let workbook = XlsxReader::read_file(&path).expect("read workbook");
    let mut out = Vec::new();
    XlsxWriter::write(&workbook, Cursor::new(&mut out)).expect("write workbook");

    let workbook2 = XlsxReader::read(Cursor::new(&out)).expect("read roundtripped workbook");
    assert_eq!(workbook2.sheet_count(), 1);

    let styles_xml = read_styles_xml_from_xlsx(&out);
    let snapshot = parse_styles_snapshot(&styles_xml);

    assert_eq!(snapshot.cell_style_xf_count, 3);
    assert!(snapshot.named_styles.iter().any(|name| name == "Normal"));
    assert!(snapshot.named_styles.iter().any(|name| name == "Heading 1"));
    assert!(snapshot.named_styles.iter().any(|name| name == "Currency"));
    assert!(snapshot.cell_xf_ids.iter().any(|xf_id| *xf_id == 1));

    cleanup_fixture(&path);
}
