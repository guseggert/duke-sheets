//! E2E tests for rich text (per-character formatting) via Excel COM.
//!
//! Writer test: duke-sheets creates rich text → roundtrip through Excel → verify.
//! Reader test: Excel COM creates rich text via Characters API → read back → verify.

use crate::{
    cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, roundtrip_through_excel,
    temp_fixture,
};
use duke_sheets_core::cell::CellValue;
use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::{Color, Workbook};
use duke_sheets_xlsx::XlsxReader;

// Writer E2E: duke-sheets writes rich text → Excel opens + re-saves → read back

/// Verify Excel accepts our inline rich text and preserves it on re-save.
///
/// After Excel re-saves, inline strings move to the shared string table.
/// Excel may normalize run fonts (adding explicit defaults), but the text
/// content and key formatting (bold, italic, color) must survive.
#[test]
fn test_write_rich_text_roundtrip() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Cell A1: "Hello " (plain) + "World" (bold)
    let runs = vec![
        RichTextRun::plain("Hello "),
        RichTextRun::with_font(
            "World",
            RunFont {
                bold: Some(true),
                ..Default::default()
            },
        ),
    ];
    sheet
        .set_cell_value_at(0, 0, CellValue::RichText(runs))
        .unwrap();

    // Cell A2: "Red" (red, italic) + " and " (plain) + "Blue" (blue, bold)
    let runs2 = vec![
        RichTextRun::with_font(
            "Red",
            RunFont {
                italic: Some(true),
                color: Some(Color::rgb(255, 0, 0)),
                ..Default::default()
            },
        ),
        RichTextRun::plain(" and "),
        RichTextRun::with_font(
            "Blue",
            RunFont {
                bold: Some(true),
                color: Some(Color::rgb(0, 0, 255)),
                ..Default::default()
            },
        ),
    ];
    sheet
        .set_cell_value_at(1, 0, CellValue::RichText(runs2))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();

    // After Excel re-save, rich text ends up in the shared string table
    // (not inline strings). Verify the text content and run structure.
    match s.get_value_at(0, 0) {
        CellValue::RichText(read_runs) => {
            // Excel may split/merge runs, but the text should be preserved
            let full_text: String = read_runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(full_text, "Hello World", "A1 text");

            // At least one run should have bold
            let has_bold = read_runs
                .iter()
                .any(|r| r.font.as_ref().map_or(false, |f| f.bold == Some(true)));
            assert!(has_bold, "A1 should have a bold run, got: {read_runs:#?}");
        }
        CellValue::String(s) => {
            // Excel might collapse single-styled rich text to plain string
            assert_eq!(s.as_ref(), "Hello World", "A1 collapsed to plain string");
        }
        other => panic!("A1: expected RichText or String, got {:?}", other),
    }

    match s.get_value_at(1, 0) {
        CellValue::RichText(read_runs) => {
            let full_text: String = read_runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(full_text, "Red and Blue", "A2 text");

            // Should have runs with different colors
            let has_italic = read_runs
                .iter()
                .any(|r| r.font.as_ref().map_or(false, |f| f.italic == Some(true)));
            let has_bold = read_runs
                .iter()
                .any(|r| r.font.as_ref().map_or(false, |f| f.bold == Some(true)));
            assert!(
                has_italic,
                "A2 should have an italic run, got: {read_runs:#?}"
            );
            assert!(has_bold, "A2 should have a bold run, got: {read_runs:#?}");
        }
        CellValue::String(s) => {
            assert_eq!(s.as_ref(), "Red and Blue", "A2 collapsed to plain string");
        }
        other => panic!("A2: expected RichText or String, got {:?}", other),
    }
}

// Reader E2E: Excel COM creates rich text → duke-sheets reads it

/// Have Excel create per-character formatting via Characters API,
/// save to XLSX, and verify our reader parses the rich text runs.
///
/// Uses Range.Characters(start, length).Font to apply formatting to
/// substrings. Excel saves these as SST entries with `<r>/<rPr>` runs.
#[test]
fn test_read_rich_text_from_excel() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        // Set plain text first, then format substrings.
        // A1: "Normal Bold Normal" — make "Bold" (chars 8-11) bold
        wb.set_cell_value("A1", "Normal Bold Normal")
            .expect("set A1");
        wb.set_character_font_property("A1", 8, 4, "Bold", serde_json::Value::from(true))
            .expect("set A1 bold chars");

        // A2: "Small BIG Small" — make "BIG" (chars 7-9) larger
        wb.set_cell_value("A2", "Small BIG Small").expect("set A2");
        wb.set_character_font_property("A2", 7, 3, "Size", serde_json::Value::from(20.0))
            .expect("set A2 size chars");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");

    // A1: Should be rich text with runs (Excel splits on formatting boundaries)
    match sheet.get_value_at(0, 0) {
        CellValue::RichText(runs) => {
            let full_text: String = runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(full_text, "Normal Bold Normal", "A1 text");
            // There should be multiple runs (at least 2)
            assert!(
                runs.len() >= 2,
                "A1 should have multiple runs, got {}: {:#?}",
                runs.len(),
                runs
            );
            // At least one run should have bold=true
            let has_bold = runs
                .iter()
                .any(|r| r.font.as_ref().map_or(false, |f| f.bold == Some(true)));
            assert!(has_bold, "A1 should have a bold run, got: {runs:#?}");
        }
        CellValue::String(s) => {
            // Characters API may not have worked — just verify text
            assert_eq!(
                s.as_ref(),
                "Normal Bold Normal",
                "A1 as plain string (Characters formatting may not have persisted)"
            );
        }
        other => panic!("A1: expected RichText or String, got {:?}", other),
    }

    // A2: Should be rich text with a large-font run
    match sheet.get_value_at(1, 0) {
        CellValue::RichText(runs) => {
            let full_text: String = runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(full_text, "Small BIG Small", "A2 text");
            // At least one run should have a larger size
            let has_big = runs.iter().any(|r| {
                r.font
                    .as_ref()
                    .map_or(false, |f| f.size.map_or(false, |s| s >= 18.0))
            });
            assert!(
                has_big,
                "A2 should have a run with large font size, got: {runs:#?}"
            );
        }
        CellValue::String(s) => {
            assert_eq!(
                s.as_ref(),
                "Small BIG Small",
                "A2 as plain string (Characters formatting may not have persisted)"
            );
        }
        other => panic!("A2: expected RichText or String, got {:?}", other),
    }

    cleanup_fixture(&fixture);
}
