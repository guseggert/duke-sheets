use std::io::{Cursor, Read, Write};

use duke_sheets_core::CellValue;
use duke_sheets_xlsx::{XlsxDiagnosticCode, XlsxPackagePolicy, XlsxReader};

fn relocated_package(root_target: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    let root_relationships = format!(
        r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{root_target}"/></Relationships>"#
    );
    let parts: [(&str, &[u8]); 8] = [
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/pkg/main/book.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/sheets/data.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/pkg/resources/strings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/pkg/resources/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/pkg/theme/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            root_relationships.as_bytes(),
        ),
        (
            "pkg/main/book.xml",
            br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Moved" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "pkg/main/_rels/book.xml.rels",
            br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../../sheets/data.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="../resources/strings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../resources/styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/custom.xml"/></Relationships>"#,
        ),
        (
            "sheets/data.xml",
            br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1" t="s"><v>0</v></c></row></sheetData></worksheet>"#,
        ),
        (
            "pkg/resources/strings.xml",
            br#"<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>relocated</t></si></sst>"#,
        ),
        (
            "pkg/resources/styles.xml",
            br#"<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>"#,
        ),
        (
            "pkg/theme/custom.xml",
            br#"<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Moved"><a:themeElements><a:clrScheme name="Moved"><a:dk1><a:srgbClr val="010101"/></a:dk1><a:lt1><a:srgbClr val="020202"/></a:lt1><a:dk2><a:srgbClr val="030303"/></a:dk2><a:lt2><a:srgbClr val="040404"/></a:lt2><a:accent1><a:srgbClr val="112233"/></a:accent1><a:accent2><a:srgbClr val="060606"/></a:accent2><a:accent3><a:srgbClr val="070707"/></a:accent3><a:accent4><a:srgbClr val="080808"/></a:accent4><a:accent5><a:srgbClr val="090909"/></a:accent5><a:accent6><a:srgbClr val="0A0A0A"/></a:accent6><a:hlink><a:srgbClr val="0B0B0B"/></a:hlink><a:folHlink><a:srgbClr val="0C0C0C"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>"#,
        ),
    ];

    for (name, bytes) in parts {
        zip.start_file(name, options).expect("start part");
        zip.write_all(bytes).expect("write part");
    }
    zip.finish().expect("finish package").into_inner()
}

fn assert_relocated_package(bytes: Vec<u8>) {
    let workbook = XlsxReader::read(Cursor::new(bytes)).expect("read relocated package");
    let sheet = workbook.worksheet(0).expect("worksheet");
    assert_eq!(sheet.name(), "Moved");
    assert_eq!(
        sheet.cell_at(0, 0).map(|cell| &cell.value),
        Some(&CellValue::String("relocated".into()))
    );
    assert!(sheet.cell_style_at(0, 0).expect("cell style").font.bold);
    assert_eq!(workbook.theme_palette().theme_rgb(4), (0x11, 0x22, 0x33));
}

fn without_part(bytes: Vec<u8>, omitted: &str) -> Vec<u8> {
    let mut source = zip::ZipArchive::new(Cursor::new(bytes)).expect("open package");
    let mut target = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let mut file = source.by_index(index).expect("source part");
        if file.name() == omitted {
            continue;
        }
        let name = file.name().to_string();
        let mut contents = Vec::new();
        file.read_to_end(&mut contents).expect("read source part");
        target
            .start_file(name, zip::write::SimpleFileOptions::default())
            .expect("start copied part");
        target.write_all(&contents).expect("write copied part");
    }
    target.finish().expect("finish copied package").into_inner()
}

fn rename_parts(bytes: Vec<u8>, renames: &[(&str, &str)]) -> Vec<u8> {
    let mut source = zip::ZipArchive::new(Cursor::new(bytes)).expect("open package");
    let mut target = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let mut file = source.by_index(index).expect("source part");
        let name = renames
            .iter()
            .find_map(|(from, to)| (file.name() == *from).then_some(*to))
            .unwrap_or(file.name())
            .to_string();
        let mut contents = Vec::new();
        file.read_to_end(&mut contents).expect("read source part");
        target
            .start_file(name, zip::write::SimpleFileOptions::default())
            .expect("start renamed part");
        target.write_all(&contents).expect("write renamed part");
    }
    target
        .finish()
        .expect("finish renamed package")
        .into_inner()
}

fn rewrite_text_part(
    bytes: Vec<u8>,
    rewritten: &str,
    edit: impl FnOnce(String) -> String,
) -> Vec<u8> {
    let mut source = zip::ZipArchive::new(Cursor::new(bytes)).expect("open package");
    let mut target = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut edit = Some(edit);
    for index in 0..source.len() {
        let mut file = source.by_index(index).expect("source part");
        let name = file.name().to_string();
        let mut contents = Vec::new();
        file.read_to_end(&mut contents).expect("read source part");
        if name == rewritten {
            let text = String::from_utf8(contents).expect("text part");
            contents = edit.take().expect("single matching part")(text).into_bytes();
        }
        target
            .start_file(name, zip::write::SimpleFileOptions::default())
            .expect("start copied part");
        target.write_all(&contents).expect("write copied part");
    }
    target.finish().expect("finish copied package").into_inner()
}

#[test]
fn reads_relocated_workbook_and_resources_from_relative_relationships() {
    assert_relocated_package(relocated_package("pkg/main/book.xml"));
}

#[test]
fn reads_relocated_workbook_from_root_relative_relationship() {
    assert_relocated_package(relocated_package("/pkg/main/book.xml"));
}

#[test]
fn compatible_finds_conventional_resources_without_relationships() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "pkg/main/_rels/book.xml.rels",
        |relationships| {
            relationships
                .replace(r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="../resources/strings.xml"/>"#, "")
                .replace(r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../resources/styles.xml"/>"#, "")
                .replace(r#"<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/custom.xml"/>"#, "")
        },
    );
    let bytes = rewrite_text_part(bytes, "[Content_Types].xml", |content_types| {
        content_types
            .replace("/pkg/resources/strings.xml", "/pkg/main/sharedStrings.xml")
            .replace("/pkg/resources/styles.xml", "/pkg/main/styles.xml")
            .replace("/pkg/theme/custom.xml", "/pkg/main/theme/theme1.xml")
    });
    let bytes = rename_parts(
        bytes,
        &[
            ("pkg/resources/strings.xml", "pkg/main/sharedStrings.xml"),
            ("pkg/resources/styles.xml", "pkg/main/styles.xml"),
            ("pkg/theme/custom.xml", "pkg/main/theme/theme1.xml"),
        ],
    );
    assert_relocated_package(bytes);
}

#[test]
fn strict_mode_accepts_a_valid_relocated_package() {
    let report = XlsxReader::read_with_report(
        Cursor::new(relocated_package("pkg/main/book.xml")),
        XlsxPackagePolicy::Strict,
    )
    .expect("strict read");
    assert_eq!(
        report.workbook.worksheet(0).expect("worksheet").name(),
        "Moved"
    );
    assert!(report.diagnostics.is_empty());
}

// features: Package validation diagnostics
#[test]
fn compatible_mode_recovers_workbook_from_content_types_without_root_rels() {
    let bytes = without_part(relocated_package("pkg/main/book.xml"), "_rels/.rels");
    let report = XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Compatible)
        .expect("compatible read");
    assert_eq!(
        report.workbook.worksheet(0).expect("worksheet").name(),
        "Moved"
    );
    assert!(report
        .diagnostics
        .iter()
        .any(|diagnostic| { diagnostic.code == XlsxDiagnosticCode::MissingPackageRelationships }));
    assert!(report
        .diagnostics
        .iter()
        .any(|diagnostic| { diagnostic.code == XlsxDiagnosticCode::CanonicalPartFallback }));
}

// features: Package validation diagnostics
#[test]
fn strict_mode_rejects_package_without_root_rels() {
    let bytes = without_part(relocated_package("pkg/main/book.xml"), "_rels/.rels");
    assert!(XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict).is_err());
}

#[test]
fn strict_mode_rejects_external_worksheet_relationship() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "pkg/main/_rels/book.xml.rels",
        |rels| {
            rels.replace(
                "Target=\"../../sheets/data.xml\"",
                "Target=\"https://example.com/data.xml\" TargetMode=\"External\"",
            )
        },
    );
    assert!(XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict).is_err());
}

#[test]
fn strict_mode_rejects_wrong_worksheet_content_type() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "[Content_Types].xml",
        |content_types| {
            content_types.replace(
                r#"<Override PartName="/sheets/data.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                "",
            )
        },
    );
    assert!(XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict).is_err());
}

#[test]
fn compatible_mode_recovers_when_root_rels_omit_office_document() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "_rels/.rels",
        |relationships| {
            relationships.replace(
                r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="pkg/main/book.xml"/>"#,
                "",
            )
        },
    );
    let report = XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Compatible)
        .expect("compatible read");
    assert_eq!(
        report.workbook.worksheet(0).expect("worksheet").name(),
        "Moved"
    );
    assert!(report.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == XlsxDiagnosticCode::MissingOfficeDocumentRelationship
    }));
}

#[test]
fn compatible_accepts_nonstandard_workbook_root_but_strict_rejects_it() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "pkg/main/book.xml",
        |workbook| {
            workbook
                .replace("<workbook ", "<notWorkbook ")
                .replace("</workbook>", "</notWorkbook>")
        },
    );
    assert!(XlsxReader::read(Cursor::new(&bytes)).is_ok());
    assert!(XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict,).is_err());
}

#[test]
fn compatible_falls_back_from_missing_office_document_target() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "_rels/.rels",
        |relationships| {
            relationships.replace(r#"Target="pkg/main/book.xml""#, r#"Target="missing.xml""#)
        },
    );
    let report = XlsxReader::read_with_report(Cursor::new(&bytes), XlsxPackagePolicy::Compatible)
        .expect("compatible fallback");
    assert_eq!(
        report.workbook.worksheet(0).expect("worksheet").name(),
        "Moved"
    );
    assert!(report
        .diagnostics
        .iter()
        .any(|diagnostic| { diagnostic.code == XlsxDiagnosticCode::MissingRelationshipTarget }));
    assert!(XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict,).is_err());
}

#[test]
fn compatible_falls_back_from_malformed_package_relationships() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "_rels/.rels",
        |relationships| relationships.replace("</Relationships>", "<broken>"),
    );
    let report = XlsxReader::read_with_report(Cursor::new(&bytes), XlsxPackagePolicy::Compatible)
        .expect("compatible fallback");
    assert_eq!(
        report.workbook.worksheet(0).expect("worksheet").name(),
        "Moved"
    );
    assert!(report
        .diagnostics
        .iter()
        .any(|diagnostic| { diagnostic.code == XlsxDiagnosticCode::MalformedRelationship }));
    assert!(XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict).is_err());
}

#[test]
fn compatible_prefers_workbook_content_type_among_ambiguous_roots() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "_rels/.rels",
        |relationships| {
            relationships.replace(
                "<Relationship Id=\"rId1\"",
                "<Relationship Id=\"rId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"pkg/resources/styles.xml\"/><Relationship Id=\"rId1\"",
            )
        },
    );
    let report = XlsxReader::read_with_report(Cursor::new(&bytes), XlsxPackagePolicy::Compatible)
        .expect("compatible selection");
    assert_eq!(
        report.workbook.worksheet(0).expect("worksheet").name(),
        "Moved"
    );
    assert!(report.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == XlsxDiagnosticCode::AmbiguousOfficeDocumentRelationship
    }));
    assert!(XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict).is_err());
}

#[test]
fn compatible_ignores_malformed_content_types_but_strict_rejects_them() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "[Content_Types].xml",
        |content_types| content_types.replace("</Types>", "<broken>"),
    );
    let report = XlsxReader::read_with_report(Cursor::new(&bytes), XlsxPackagePolicy::Compatible)
        .expect("compatible read");
    assert_eq!(
        report.workbook.worksheet(0).expect("worksheet").name(),
        "Moved"
    );
    assert!(report
        .diagnostics
        .iter()
        .any(|diagnostic| { diagnostic.code == XlsxDiagnosticCode::MalformedContentType }));
    assert!(XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict,).is_err());
}

#[test]
fn strict_skips_valid_unmodeled_dialog_sheet_without_inventing_a_worksheet() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "pkg/main/_rels/book.xml.rels",
        |relationships| {
            relationships.replace(
                r#"/relationships/worksheet""#,
                r#"/relationships/dialogsheet""#,
            )
        },
    );
    let bytes = rewrite_text_part(bytes, "[Content_Types].xml", |content_types| {
        content_types.replace(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
        )
    });
    let report = XlsxReader::read_with_report(Cursor::new(bytes), XlsxPackagePolicy::Strict)
        .expect("strict read");
    assert_eq!(report.workbook.sheet_count(), 0);
    assert!(report
        .diagnostics
        .iter()
        .any(|diagnostic| { diagnostic.code == XlsxDiagnosticCode::UnsupportedSheetType }));
}

#[test]
fn accepts_processing_instruction_before_workbook_root() {
    let bytes = rewrite_text_part(
        relocated_package("pkg/main/book.xml"),
        "pkg/main/book.xml",
        |workbook| workbook.replacen("?>", "?><?probe compatible?>", 1),
    );
    assert!(XlsxReader::read(Cursor::new(bytes)).is_ok());
}
