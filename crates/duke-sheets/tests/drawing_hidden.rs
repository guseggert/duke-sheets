//! Drawing-object hidden flag (`DrawingMeta.hidden`) round-trip
//! coverage for non-comment kinds.
//!
//! XLSX/XLSB carry it on the native shape's `cNvPr@hidden` (images,
//! charts, groups) and on the legacy VML shape's `visibility:hidden`
//! style token (form controls). XLS packs it into the Escher FOPT
//! Group Shape Boolean Properties entry (opid 0x03BF, fHidden).
//! Absent markup always means visible.

use std::io::Cursor;

use duke_sheets::{
    CellMarker, ChildTransform, DrawingAnchor, DrawingKind, DrawingMeta, DrawingObject,
    EmbeddedImage, FormControl, FormControlKind, Group, GroupChild, GroupTransform, ImageFormat,
    Workbook, Worksheet,
};
use duke_sheets_xls::{XlsReader, XlsWriter};
use duke_sheets_xlsb::{XlsbReader, XlsbWriter};
use duke_sheets_xlsx::{XlsxReader, XlsxWriter};

const PNG_1PX: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
];

fn two_cell(from_col: u16, from_row: u32, to_col: u16, to_row: u32) -> DrawingAnchor {
    DrawingAnchor::TwoCell {
        from: CellMarker {
            col: from_col,
            col_offset_emu: 0,
            row: from_row,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: to_col,
            col_offset_emu: 0,
            row: to_row,
            row_offset_emu: 0,
        },
        edit_as: None,
    }
}

fn png(name: &str) -> DrawingObject {
    DrawingObject::image(EmbeddedImage {
        format: ImageFormat::Png,
        media_path: String::new(),
        svg_media_path: None,
        width_emu: 190_500,
        height_emu: 190_500,
        rotation: None,
        flip_h: false,
        flip_v: false,
        data: PNG_1PX.to_vec(),
        svg_data: None,
    })
    .with_name(name)
}

fn checkbox(caption: &str) -> FormControl {
    FormControl::new(FormControlKind::Checkbox {
        caption: caption.into(),
        state: duke_sheets::CheckState::Checked,
        cell_link: None,
        no_3d: false,
    })
}

fn round_trip_xlsx(workbook: &Workbook) -> Workbook {
    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(workbook, &mut output).expect("write xlsx");
    XlsxReader::read(Cursor::new(output.into_inner())).expect("read xlsx")
}

fn round_trip_xlsb(workbook: &Workbook) -> Workbook {
    let mut output = Cursor::new(Vec::new());
    XlsbWriter::write(workbook, &mut output).expect("write xlsb");
    XlsbReader::read(Cursor::new(output.into_inner())).expect("read xlsb")
}

fn round_trip_xls(workbook: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(workbook).expect("write xls");
    XlsReader::read(Cursor::new(bytes)).expect("read xls")
}

/// One sheet carrying a visible image, a hidden image, a visible
/// control, and a hidden control.
fn build_mixed_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Shown").with_anchor(two_cell(0, 0, 2, 2)));
    sheet.add_drawing(
        png("Ghost")
            .with_anchor(two_cell(2, 2, 4, 4))
            .with_hidden(true),
    );
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Visible box")).with_anchor(two_cell(4, 4, 6, 6)),
    );
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Cloaked box"))
            .with_anchor(two_cell(6, 6, 8, 8))
            .with_hidden(true),
    );
    workbook
}

/// Hidden survives on the flagged image/control, and stays false on
/// the default ones.
fn assert_mixed_hidden_flags(sheet: &Worksheet, format: &str) {
    let images: Vec<_> = sheet.images().collect();
    assert_eq!(images.len(), 2, "{format}: both images survive");
    assert_eq!(images[0].object.meta.name.as_deref(), Some("Shown"));
    assert!(
        !images[0].object.meta.hidden,
        "{format}: default image must read back hidden == false"
    );
    assert_eq!(images[1].object.meta.name.as_deref(), Some("Ghost"));
    assert!(
        images[1].object.meta.hidden,
        "{format}: hidden image must read back hidden == true"
    );

    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), 2, "{format}: both controls survive");
    assert_eq!(
        controls[0].payload.caption_text().as_deref(),
        Some("Visible box")
    );
    assert!(
        !controls[0].object.meta.hidden,
        "{format}: default control must read back hidden == false"
    );
    assert_eq!(
        controls[1].payload.caption_text().as_deref(),
        Some("Cloaked box")
    );
    assert!(
        controls[1].object.meta.hidden,
        "{format}: hidden control must read back hidden == true"
    );
}

#[test]
fn xlsx_hidden_image_and_control_round_trip() {
    let read = round_trip_xlsx(&build_mixed_workbook());
    assert_mixed_hidden_flags(read.worksheet(0).unwrap(), "xlsx");
}

#[test]
fn xlsb_hidden_image_and_control_round_trip() {
    let read = round_trip_xlsb(&build_mixed_workbook());
    assert_mixed_hidden_flags(read.worksheet(0).unwrap(), "xlsb");
}

#[test]
fn xls_hidden_image_and_control_round_trip() {
    let read = round_trip_xls(&build_mixed_workbook());
    assert_mixed_hidden_flags(read.worksheet(0).unwrap(), "xls");
}

/// A hidden image inside a group keeps its child meta.hidden through
/// an XLSX round trip; the sibling stays visible.
#[test]
fn xlsx_hidden_group_child_round_trips() {
    let child = |name: &str, x: i64, hidden: bool| GroupChild {
        meta: DrawingMeta {
            name: Some(name.to_string()),
            hidden,
            ..Default::default()
        },
        transform: ChildTransform {
            x_emu: x,
            y_emu: 0,
            cx_emu: 190_500,
            cy_emu: 190_500,
            rotation: 0,
            flip_h: false,
            flip_v: false,
        },
        kind: match png(name).kind {
            DrawingKind::Image(image) => DrawingKind::Image(image),
            _ => unreachable!(),
        },
    };
    let group = Group {
        transform: GroupTransform {
            x_emu: 609_600,
            y_emu: 190_500,
            cx_emu: 1_219_200,
            cy_emu: 190_500,
            child_x_emu: 0,
            child_y_emu: 0,
            child_cx_emu: 1_219_200,
            child_cy_emu: 190_500,
            rotation: 0,
            flip_h: false,
            flip_v: false,
        },
        children: vec![child("Left", 0, false), child("Right", 609_600, true)],
    };

    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(
        DrawingObject::group(group)
            .with_anchor(two_cell(1, 1, 4, 2))
            .with_name("Group 1"),
    );

    let read = round_trip_xlsx(&workbook);
    let object = &read.worksheet(0).unwrap().drawings()[0];
    assert!(
        !object.meta.hidden,
        "group wrapper itself stays visible by default"
    );
    let group = object.kind.as_group().expect("group survives");
    assert_eq!(group.children.len(), 2);
    assert_eq!(group.children[0].meta.name.as_deref(), Some("Left"));
    assert!(
        !group.children[0].meta.hidden,
        "default group child must read back hidden == false"
    );
    assert_eq!(group.children[1].meta.name.as_deref(), Some("Right"));
    assert!(
        group.children[1].meta.hidden,
        "hidden group child must read back hidden == true"
    );
}
