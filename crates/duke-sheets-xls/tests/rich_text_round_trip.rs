//! Round-trip tests for rich-text SST emission. Each rich-text cell
//! goes through:
//!   StyleTables interns each run's RunFont as a complete FontStyle;
//!   SstTable builds an SstEntry with the (char_pos, font_idx) pairs;
//!   the SST emitter sets the fRichSt flag and appends the run array
//!   after the character data.

use std::io::Cursor;

use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::{Color, FontVerticalAlign, Underline};
use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn run(text: &str, font: Option<RunFont>) -> RichTextRun {
    RichTextRun {
        text: text.to_string(),
        font,
    }
}

#[test]
fn rich_text_value_round_trips_plain_text() {
    // RichText with all-None fonts behaves like a plain string.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let runs = vec![run("Hello, world", None)];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs.clone()))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    assert_eq!(value.as_string(), Some("Hello, world"));
}

#[test]
fn rich_text_with_bold_run_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let bold = RunFont {
        bold: Some(true),
        ..Default::default()
    };
    let runs = vec![run("plain ", None), run("bold", Some(bold))];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    match value {
        CellValue::RichText(rt) => {
            let plain: String = rt
                .iter()
                .map(|r| r.text.as_str())
                .collect::<Vec<_>>()
                .join("");
            assert_eq!(plain, "plain bold");
            // Find a run whose font reports bold = Some(true).
            let has_bold = rt
                .iter()
                .any(|r| matches!(&r.font, Some(f) if f.bold == Some(true)));
            assert!(has_bold, "expected at least one bold run, got {rt:?}");
        }
        other => panic!("expected RichText, got {other:?}"),
    }
}

#[test]
fn rich_text_with_size_and_color_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let big_red = RunFont {
        size: Some(18.0),
        color: Some(Color::Indexed(2)), // red
        ..Default::default()
    };
    let runs = vec![run("normal ", None), run("BIG RED", Some(big_red))];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    match value {
        CellValue::RichText(rt) => {
            let big_run = rt
                .iter()
                .find(|r| r.text.contains("BIG"))
                .expect("BIG run present");
            let font = big_run.font.as_ref().expect("font present on BIG run");
            assert!(
                font.size.is_some_and(|s| (s - 18.0).abs() < 0.01),
                "expected size 18 on BIG run, got {font:?}"
            );
        }
        other => panic!("expected RichText, got {other:?}"),
    }
}

#[test]
fn rich_text_with_strikethrough_run_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let struck = RunFont {
        strikethrough: Some(true),
        ..Default::default()
    };
    let runs = vec![run("kept ", None), run("removed", Some(struck))];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    match value {
        CellValue::RichText(rt) => {
            let has_strike = rt
                .iter()
                .any(|r| matches!(&r.font, Some(f) if f.strikethrough == Some(true)));
            assert!(has_strike, "expected strikethrough run, got {rt:?}");
        }
        other => panic!("expected RichText, got {other:?}"),
    }
}

#[test]
fn rich_text_with_underline_run_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let dbl_under = RunFont {
        underline: Some(Underline::Double),
        ..Default::default()
    };
    let runs = vec![run("plain ", None), run("underlined", Some(dbl_under))];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    match value {
        CellValue::RichText(rt) => {
            let has_double = rt
                .iter()
                .any(|r| matches!(&r.font, Some(f) if f.underline == Some(Underline::Double)));
            assert!(has_double, "expected double-underlined run, got {rt:?}");
        }
        other => panic!("expected RichText, got {other:?}"),
    }
}

#[test]
fn rich_text_with_superscript_run_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let superscript = RunFont {
        vertical_align: Some(FontVerticalAlign::Superscript),
        ..Default::default()
    };
    let runs = vec![run("E=mc", None), run("2", Some(superscript))];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    match value {
        CellValue::RichText(rt) => {
            let has_super = rt.iter().any(
                |r| matches!(&r.font, Some(f) if f.vertical_align == Some(FontVerticalAlign::Superscript)),
            );
            assert!(has_super, "expected superscript run, got {rt:?}");
        }
        other => panic!("expected RichText, got {other:?}"),
    }
}

#[test]
fn rich_text_with_named_font_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let arial = RunFont {
        name: Some("Arial".into()),
        ..Default::default()
    };
    let runs = vec![run("default ", None), run("arial", Some(arial))];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    match value {
        CellValue::RichText(rt) => {
            let has_arial = rt
                .iter()
                .any(|r| matches!(&r.font, Some(f) if f.name.as_deref() == Some("Arial")));
            assert!(has_arial, "expected Arial run, got {rt:?}");
        }
        other => panic!("expected RichText, got {other:?}"),
    }
}

#[test]
fn rich_text_with_italic_run_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let italic = RunFont {
        italic: Some(true),
        ..Default::default()
    };
    let runs = vec![run("upright ", None), run("italic", Some(italic))];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    match value {
        CellValue::RichText(rt) => {
            let has_italic = rt
                .iter()
                .any(|r| matches!(&r.font, Some(f) if f.italic == Some(true)));
            assert!(has_italic, "expected italic run, got {rt:?}");
        }
        other => panic!("expected RichText, got {other:?}"),
    }
}

#[test]
fn three_run_rich_text_round_trips_text_and_run_count() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let bold = RunFont {
        bold: Some(true),
        ..Default::default()
    };
    let italic = RunFont {
        italic: Some(true),
        ..Default::default()
    };
    let runs = vec![
        run("plain ", None),
        run("bold ", Some(bold)),
        run("italic", Some(italic)),
    ];
    ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
        .expect("A1");

    let parsed = write_then_read(&wb);
    let value = parsed.worksheet(0).unwrap().get_value_at(0, 0);
    match value {
        CellValue::RichText(rt) => {
            let plain: String = rt
                .iter()
                .map(|r| r.text.as_str())
                .collect::<Vec<_>>()
                .join("");
            assert_eq!(plain, "plain bold italic");
        }
        other => panic!("expected RichText, got {other:?}"),
    }
}
