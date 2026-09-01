use std::io::{Cursor, Read, Write};

use duke_sheets::{Color, Workbook, XlsbReader, XlsbWriter, XlsxReader, XlsxWriter};

/// A theme part whose clrScheme colors slot `i` as `(10+i, 20+i, 30+i)`,
/// covering both `srgbClr` and `sysClr`/`lastClr` carriers.
const THEME_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture">
 <a:themeElements>
  <a:clrScheme name="Fixture">
   <a:dk1><a:sysClr val="windowText" lastClr="0B151F"/></a:dk1>
   <a:lt1><a:sysClr val="window" lastClr="0A141E"/></a:lt1>
   <a:dk2><a:srgbClr val="0D1721"/></a:dk2>
   <a:lt2><a:srgbClr val="0C1620"/></a:lt2>
   <a:accent1><a:srgbClr val="0E1822"/></a:accent1>
   <a:accent2><a:srgbClr val="0F1923"/></a:accent2>
   <a:accent3><a:srgbClr val="101A24"/></a:accent3>
   <a:accent4><a:srgbClr val="111B25"/></a:accent4>
   <a:accent5><a:srgbClr val="121C26"/></a:accent5>
   <a:accent6><a:srgbClr val="131D27"/></a:accent6>
   <a:hlink><a:srgbClr val="141E28"/></a:hlink>
   <a:folHlink><a:srgbClr val="151F29"/></a:folHlink>
  </a:clrScheme>
 </a:themeElements>
</a:theme>"#;

fn expected_slot(i: usize) -> (u8, u8, u8) {
    (10 + i as u8, 20 + i as u8, 30 + i as u8)
}

/// Replace or add `xl/theme/theme1.xml` in an OOXML package.
fn with_theme_part(bytes: Vec<u8>) -> Vec<u8> {
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..input.len() {
        let mut file = input.by_index(index).expect("zip entry");
        let name = file.name().to_string();
        if name == "xl/theme/theme1.xml" {
            continue;
        }
        let mut data = Vec::new();
        file.read_to_end(&mut data).expect("read zip entry");
        drop(file);
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .expect("copy file");
        output.write_all(&data).expect("write zip entry");
    }
    output
        .start_file(
            "xl/theme/theme1.xml",
            zip::write::SimpleFileOptions::default(),
        )
        .expect("add theme part");
    output
        .write_all(THEME_XML.as_bytes())
        .expect("write theme part");
    output.finish().expect("finish zip").into_inner()
}

fn assert_fixture_palette(workbook: &Workbook, label: &str) {
    let palette = workbook.theme_palette();
    for (i, &color) in palette.colors.iter().enumerate() {
        assert_eq!(
            color,
            expected_slot(i),
            "{label}: clrScheme slot {i} must come from the file"
        );
    }
    assert_eq!(
        workbook.resolve_color(&Color::theme(4, 0.0)),
        Some(expected_slot(4)),
        "{label}: resolve_color must use the stored palette"
    );
    assert_eq!(workbook.resolve_color(&Color::Auto), None);
}

// features: Theme color scheme (12 slots)
#[test]
fn xlsx_reader_stores_the_file_theme_palette() {
    let mut bytes = Cursor::new(Vec::new());
    XlsxWriter::write(&Workbook::new(), &mut bytes).expect("write xlsx");
    let themed = with_theme_part(bytes.into_inner());

    let reopened = XlsxReader::read(Cursor::new(themed)).expect("read themed xlsx");
    assert_fixture_palette(&reopened, "xlsx read");

    // The raw theme part round-trips, so a re-save keeps the palette.
    let mut resaved = Cursor::new(Vec::new());
    XlsxWriter::write(&reopened, &mut resaved).expect("rewrite xlsx");
    let reread = XlsxReader::read(Cursor::new(resaved.into_inner())).expect("reread xlsx");
    assert_fixture_palette(&reread, "xlsx round trip");
}

// features: Theme color scheme (12 slots)
#[test]
fn xlsb_reader_stores_the_file_theme_palette() {
    let mut bytes = Cursor::new(Vec::new());
    XlsbWriter::write(&Workbook::new(), &mut bytes).expect("write xlsb");
    let themed = with_theme_part(bytes.into_inner());

    let reopened = XlsbReader::read(Cursor::new(themed)).expect("read themed xlsb");
    assert_fixture_palette(&reopened, "xlsb read");
}

#[test]
fn workbooks_without_a_theme_part_use_the_office_palette() {
    let workbook = Workbook::new();
    assert_eq!(
        workbook.theme_palette(),
        duke_sheets::ThemePalette::default()
    );
    assert_eq!(
        workbook.resolve_color(&Color::theme(4, 0.0)),
        Some((79, 129, 189)),
        "default accent 1"
    );
}
