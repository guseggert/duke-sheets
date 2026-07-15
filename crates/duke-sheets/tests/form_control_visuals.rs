use std::io::Cursor;

use duke_sheets::{
    CellMarker, CheckState, Color, ControlText, DrawingAnchor, DrawingObject, FormControl,
    FormControlKind, HorizontalAlignment, RichTextRun, RunFont, VerticalAlignment, Workbook,
};
use duke_sheets_core::style::Underline;
use duke_sheets_xls::{XlsReader, XlsWriter};
use duke_sheets_xlsb::{XlsbReader, XlsbWriter};
use duke_sheets_xlsx::{XlsxReader, XlsxWriter};

fn probe_workbook() -> Workbook {
    let text = ControlText {
        runs: vec![
            RichTextRun::with_font(
                "Red ",
                RunFont {
                    name: Some("Segoe UI".into()),
                    size: Some(9.0),
                    color: Some(Color::rgb(255, 0, 0)),
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
            RichTextRun::with_font(
                "Blue",
                RunFont {
                    name: Some("Arial".into()),
                    size: Some(12.0),
                    color: Some(Color::rgb(0, 0, 255)),
                    italic: Some(true),
                    underline: Some(Underline::Single),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: Some(HorizontalAlignment::Right),
        vertical_alignment: Some(VerticalAlignment::Bottom),
    };
    let control = FormControl::new(FormControlKind::Checkbox {
        caption: text,
        state: CheckState::Checked,
        cell_link: None,
        no_3d: false,
    })
    .with_macro_name("RunProbe");
    let mut object = DrawingObject::form_control(control).with_anchor(DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            col_offset_emu: 0,
            row: 2,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: 4,
            col_offset_emu: 0,
            row: 4,
            row_offset_emu: 0,
        },
        edit_as: None,
    });
    object.meta.name = Some("Visual Probe".into());
    object.meta.alt_text = Some("Visual probe alternative".into());
    object.meta.title = Some("Visual probe title".into());

    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();
    workbook
}

fn assert_visuals(workbook: &Workbook, title_supported: bool) {
    let drawn = workbook
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("control survives");
    assert_eq!(drawn.object.meta.name.as_deref(), Some("Visual Probe"));
    assert_eq!(
        drawn.object.meta.alt_text.as_deref(),
        Some("Visual probe alternative")
    );
    if title_supported {
        assert_eq!(
            drawn.object.meta.title.as_deref(),
            Some("Visual probe title")
        );
    }

    let control = drawn.payload;
    assert_eq!(control.macro_name.as_deref(), Some("RunProbe"));
    assert_eq!(control.caption_text().as_deref(), Some("Red Blue"));
    let text = control.caption().expect("captioned control");
    assert_eq!(text.horizontal_alignment, Some(HorizontalAlignment::Right));
    assert_eq!(text.vertical_alignment, Some(VerticalAlignment::Bottom));
    assert_eq!(text.runs.len(), 2);

    assert_eq!(text.runs[0].text, "Red ");
    let red = text.runs[0].font.as_ref().expect("red font");
    assert_eq!(red.name.as_deref(), Some("Segoe UI"));
    assert_eq!(red.size, Some(9.0));
    assert_eq!(red.color, Some(Color::rgb(255, 0, 0)));
    assert_eq!(red.bold, Some(true));

    assert_eq!(text.runs[1].text, "Blue");
    let blue = text.runs[1].font.as_ref().expect("blue font");
    assert_eq!(blue.name.as_deref(), Some("Arial"));
    assert_eq!(blue.size, Some(12.0));
    assert_eq!(blue.color, Some(Color::rgb(0, 0, 255)));
    assert_eq!(blue.italic, Some(true));
    assert_eq!(blue.underline, Some(Underline::Single));
}

#[test]
fn control_text_plain_has_no_explicit_alignment() {
    let text = ControlText::from("plain");
    assert_eq!(text.plain_text(), "plain");
    assert!(!text.is_empty());
    assert_eq!(text.runs, vec![RichTextRun::plain("plain")]);
    assert_eq!(text.horizontal_alignment, None);
    assert_eq!(text.vertical_alignment, None);

    let invalid = FormControl::new(FormControlKind::Button {
        caption: ControlText::from("plain"),
    })
    .with_macro_name("   ");
    assert!(invalid.validate().is_err());
}

#[test]
fn xlsx_control_visual_metadata_round_trips() {
    let mut bytes = Cursor::new(Vec::new());
    XlsxWriter::write(&probe_workbook(), &mut bytes).expect("write xlsx");
    let reopened = XlsxReader::read(Cursor::new(bytes.into_inner())).expect("read xlsx");
    assert_visuals(&reopened, true);
}

#[test]
fn xlsb_control_visual_metadata_round_trips() {
    let mut bytes = Cursor::new(Vec::new());
    XlsbWriter::write(&probe_workbook(), &mut bytes).expect("write xlsb");
    let reopened = XlsbReader::read(Cursor::new(bytes.into_inner())).expect("read xlsb");
    assert_visuals(&reopened, true);
}

#[test]
fn xls_control_visual_metadata_round_trips() {
    let bytes = XlsWriter::write_to_bytes(&probe_workbook()).expect("write xls");
    let reopened = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    assert_visuals(&reopened, false);
}
