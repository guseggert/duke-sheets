//! Tests for reading conditional formatting from XLSX files created by Excel.
//! Includes DXF (differential format) style tests.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_core::{FillStyle, HorizontalAlignment, NumberFormat};
use duke_sheets_xlsx::XlsxReader;
use std::io::{Read, Write};

/// Helper: create a workbook with values in B1:B5 and a CF rule.
/// Returns the fixture for reading back.
fn create_cf_workbook(
    fixture: &crate::TempFixture,
    setup_fc: impl FnOnce(&duke_sheets_excel_com::Workbook<'_>),
) {
    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    ensure_vm_temp_dir();
    let wb = excel.create_workbook().expect("create workbook");

    for (i, val) in [10.0, 30.0, 50.0, 70.0, 90.0].iter().enumerate() {
        let cell = format!("B{}", i + 1);
        wb.set_cell_value(&cell, *val).expect("set value");
    }

    setup_fc(&wb);

    wb.save(&fixture.vm_path).expect("save");
    wb.close().expect("close");
}

#[test]
fn test_cell_is_greater_than() {
    let fixture = temp_fixture();
    create_cf_workbook(&fixture, |wb| {
        // xlCellValue=1, xlGreater=5
        let fc = wb
            .add_format_condition("B1:B5", 1, 5, "50")
            .expect("add CF");
        fc.set_fill_color(0x00FF00).expect("set fill");
        fc.set_font_bold(true).expect("set bold");
    });

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty(), "Should have at least one CF rule");

    let has_cell_is = rules
        .iter()
        .any(|r| matches!(&r.rule_type, duke_sheets_core::CfRuleType::CellIs { .. }));
    assert!(has_cell_is, "Should have a CellIs rule");

    cleanup_fixture(&fixture);
}

#[test]
fn test_cf_dxf_bold_font() {
    let fixture = temp_fixture();
    create_cf_workbook(&fixture, |wb| {
        let fc = wb
            .add_format_condition("B1:B5", 1, 5, "50")
            .expect("add CF");
        fc.set_font_bold(true).expect("set bold");
    });

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let rule = &rules[0];
    let format = rule.format.as_ref().expect("Rule should have a DXF format");
    assert!(format.font.bold, "DXF font should be bold");

    cleanup_fixture(&fixture);
}

#[test]
fn test_cf_dxf_fill() {
    let fixture = temp_fixture();
    create_cf_workbook(&fixture, |wb| {
        let fc = wb
            .add_format_condition("B1:B5", 1, 5, "50")
            .expect("add CF");
        fc.set_fill_color(0x00FF00).expect("set fill");
    });

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let format = rules[0]
        .format
        .as_ref()
        .expect("Rule should have a DXF format");
    assert!(
        format.fill != FillStyle::None,
        "DXF should have non-None fill"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_cf_dxf_alignment() {
    // Excel COM's FormatCondition object doesn't expose alignment properties
    // (DISP_E_UNKNOWNNAME). The UI can set them but COM cannot. So we:
    //   1. Create the workbook + CF rule via COM (bold font, which works)
    //   2. Pull the real Excel-produced file
    //   3. Patch the DXF in xl/styles.xml to add alignment per OOXML spec
    //   4. Verify our reader parses it correctly
    let fixture = temp_fixture();
    create_cf_workbook(&fixture, |wb| {
        let fc = wb
            .add_format_condition("B1:B5", 1, 5, "50")
            .expect("add CF");
        fc.set_font_bold(true).expect("set bold");
    });

    pull_file_from_vm(&fixture);

    // Patch the DXF: inject <alignment horizontal="center" wrapText="1"/>
    inject_dxf_alignment(&fixture.host_path);

    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let format = rules[0]
        .format
        .as_ref()
        .expect("Rule should have a DXF format");
    assert_eq!(format.alignment.horizontal, HorizontalAlignment::Center);
    assert!(format.alignment.wrap_text, "DXF should have wrap_text");

    cleanup_fixture(&fixture);
}

/// Patch the first `<dxf>` in xl/styles.xml to include alignment.
fn inject_dxf_alignment(path: &std::path::Path) {
    use zip::read::ZipArchive;
    use zip::write::SimpleFileOptions;

    let file = std::fs::File::open(path).expect("open xlsx for patching");
    let mut archive = ZipArchive::new(file).expect("read zip");

    // Read all entries into memory, patching styles.xml
    let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).expect("zip entry");
        let name = entry.name().to_string();
        let mut buf = Vec::new();
        entry.read_to_end(&mut buf).expect("read zip entry");

        if name == "xl/styles.xml" {
            let xml = String::from_utf8(buf).expect("styles.xml is utf-8");
            // Insert <alignment horizontal="center" wrapText="1"/> after
            // the first <dxf> opening or after existing DXF children.
            // The DXF typically contains <font>...</font> already.
            let patched = xml.replacen(
                "</dxf>",
                "<alignment horizontal=\"center\" wrapText=\"1\"/></dxf>",
                1,
            );
            buf = patched.into_bytes();
        }

        entries.push((name, buf));
    }
    drop(archive);

    // Rewrite the zip
    let out = std::fs::File::create(path).expect("create patched xlsx");
    let mut writer = zip::ZipWriter::new(out);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for (name, data) in &entries {
        writer.start_file(name, options).expect("start zip entry");
        writer.write_all(data).expect("write zip entry");
    }
    writer.finish().expect("finish zip");
}

#[test]
fn test_cf_dxf_number_format() {
    let fixture = temp_fixture();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        for (i, val) in [0.1, 0.3, 0.5, 0.7, 0.9].iter().enumerate() {
            let cell = format!("A{}", i + 1);
            wb.set_cell_value(&cell, *val).expect("set value");
        }

        // xlCellValue=1, xlGreater=5
        let fc = wb
            .add_format_condition("A1:A5", 1, 5, "0.5")
            .expect("add CF");
        fc.set_number_format("0.00%").expect("set number format");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let format = rules[0]
        .format
        .as_ref()
        .expect("Rule should have a DXF format");
    assert!(
        format.number_format != NumberFormat::General,
        "DXF should have non-General number format"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_cf_dxf_border() {
    let fixture = temp_fixture();
    create_cf_workbook(&fixture, |wb| {
        let fc = wb
            .add_format_condition("B1:B5", 1, 5, "50")
            .expect("add CF");
        // xlContinuous=1, xlThin=2, blue
        fc.set_border_all(1, 2, 0x0000FF).expect("set border");
    });

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let rules = sheet.conditional_formats();
    assert!(!rules.is_empty());

    let format = rules[0]
        .format
        .as_ref()
        .expect("Rule should have a DXF format");
    let has_edges = format.border.left.is_some()
        || format.border.right.is_some()
        || format.border.top.is_some()
        || format.border.bottom.is_some();
    assert!(has_edges, "DXF should have border edges");

    cleanup_fixture(&fixture);
}

#[test]
fn test_cf_multiple_rules() {
    let fixture = temp_fixture();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        for (i, val) in [10.0, 30.0, 50.0, 70.0, 90.0].iter().enumerate() {
            let cell = format!("B{}", i + 1);
            wb.set_cell_value(&cell, *val).expect("set value");
        }

        // xlCellValue=1, xlGreater=5
        let fc1 = wb
            .add_format_condition("B1:B5", 1, 5, "70")
            .expect("add CF 1");
        fc1.set_fill_color(0xFF0000).expect("set red fill");

        // xlCellValue=1, xlLess=6
        let fc2 = wb
            .add_format_condition("B1:B5", 1, 6, "30")
            .expect("add CF 2");
        fc2.set_fill_color(0x00FF00).expect("set green fill");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let rules = sheet.conditional_formats();
    assert!(
        rules.len() >= 2,
        "Should have at least 2 CF rules, got {}",
        rules.len()
    );

    for (i, rule) in rules.iter().enumerate() {
        assert!(rule.format.is_some(), "Rule {i} should have a DXF format");
    }

    cleanup_fixture(&fixture);
}
