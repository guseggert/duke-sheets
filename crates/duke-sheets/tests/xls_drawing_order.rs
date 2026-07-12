//! XLS drawing z-order, grouping, and shape-name fidelity.
//!
//! BIFF8 stores every sheet shape (pictures, comment boxes, form
//! controls, groups) in one OfficeArt container whose order is
//! z-order, so the drawings list round-trips positionally with full
//! cross-kind interleaving (unlike XLSX, where comments live in a
//! separate legacy layer).

use std::io::Cursor;

use duke_sheets::{
    CellComment, CellMarker, ChildTransform, DrawingAnchor, DrawingKind, DrawingMeta,
    DrawingObject, EmbeddedImage, FormControl, FormControlKind, Group, GroupChild,
    GroupTransform, ImageFormat, Workbook,
};
use duke_sheets_xls::{XlsReader, XlsWriter};

const PNG_1PX: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
    0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
    0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00,
    0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
    0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
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
        caption: caption.to_string(),
        state: duke_sheets::CheckState::Checked,
        cell_link: None,
        no_3d: false,
    })
}

fn round_trip(workbook: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(workbook).expect("write");
    XlsReader::read(Cursor::new(bytes)).expect("read")
}

fn kind_tags(workbook: &Workbook) -> Vec<&'static str> {
    workbook
        .worksheet(0)
        .unwrap()
        .drawings()
        .iter()
        .map(|object| match &object.kind {
            DrawingKind::Image(_) => "image",
            DrawingKind::Chart(_) => "chart",
            DrawingKind::ChartEx(_) => "chartex",
            DrawingKind::FormControl(_) => "control",
            DrawingKind::Comment { .. } => "comment",
            DrawingKind::Group(_) => "group",
            DrawingKind::Raw(_) => "raw",
        })
        .collect()
}

/// The OfficeArt container order is the drawings-list order: a
/// control between two pictures stays between them.
#[test]
fn xls_control_z_order_between_images_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2)));
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Middle")).with_anchor(two_cell(1, 1, 3, 3)),
    );
    sheet.add_drawing(png("Above").with_anchor(two_cell(2, 2, 4, 4)));

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["image", "control", "image"]);
    let sheet = read.worksheet(0).unwrap();
    let control = sheet.form_controls().next().unwrap();
    assert_eq!(control.payload.caption(), Some("Middle"));
    assert_eq!(control.object.anchor, two_cell(1, 1, 3, 3));
}

/// XLS comments are shapes in the same container, so full cross-kind
/// interleaving round-trips: image, comment, control, image.
#[test]
fn xls_full_interleave_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Bottom").with_anchor(two_cell(0, 0, 2, 2)));
    sheet.add_drawing(DrawingObject::comment(
        3,
        3,
        CellComment::new("a", "note between"),
    ));
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Check")).with_anchor(two_cell(4, 4, 6, 6)),
    );
    sheet.add_drawing(png("Top").with_anchor(two_cell(6, 6, 8, 8)));

    let read = round_trip(&workbook);
    assert_eq!(
        kind_tags(&read),
        vec!["image", "comment", "control", "image"]
    );
    let sheet = read.worksheet(0).unwrap();
    assert_eq!(sheet.comment_at(3, 3).unwrap().text, "note between");
}

/// A shape group of two pictures round-trips as a Group tree through
/// the SpgrContainer, and grouped shapes do not scramble the OBJ
/// pairing of later top-level shapes (the old reader walked
/// containers in deferred order, mispairing everything after a
/// group).
#[test]
fn xls_group_of_images_round_trips_and_keeps_later_shapes_paired() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();

    let child = |name: &str, x: i64| GroupChild {
        meta: DrawingMeta {
            name: Some(name.to_string()),
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
        children: vec![child("Left", 0), child("Right", 609_600)],
    };
    sheet.add_drawing(
        DrawingObject::group(group)
            .with_anchor(two_cell(1, 1, 4, 2))
            .with_name("Group 1"),
    );
    // A control AFTER the group: the old deferred-children walk would
    // mispair its OBJ record.
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("After group")).with_anchor(two_cell(5, 5, 7, 7)),
    );

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["group", "control"]);
    let sheet = read.worksheet(0).unwrap();
    let group = sheet.drawings()[0].kind.as_group().expect("group");
    assert_eq!(group.children.len(), 2);
    assert_eq!(group.children[0].meta.name.as_deref(), Some("Left"));
    assert_eq!(group.children[1].transform.x_emu, 609_600);
    assert!(matches!(group.children[0].kind, DrawingKind::Image(_)));
    let control = sheet.form_controls().next().unwrap();
    assert_eq!(
        control.payload.caption(),
        Some("After group"),
        "control after a group must keep its own OBJ pairing"
    );
    assert_eq!(control.object.anchor, two_cell(5, 5, 7, 7));
}

/// Control shape names round-trip through the FOPT wzName property
/// (they were previously dropped in XLS).
#[test]
fn xls_control_name_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Named"))
            .with_anchor(two_cell(1, 1, 3, 3))
            .with_name("Check Box 7"),
    );

    let read = round_trip(&workbook);
    let sheet = read.worksheet(0).unwrap();
    let control = sheet.form_controls().next().expect("control");
    assert_eq!(control.object.meta.name.as_deref(), Some("Check Box 7"));
}

/// Image alt text round-trips through FOPT wzDescription.
#[test]
fn xls_image_alt_text_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let mut object = png("Pic").with_anchor(two_cell(0, 0, 2, 2));
    object.meta.alt_text = Some("a tiny transparent pixel".to_string());
    sheet.add_drawing(object);

    let read = round_trip(&workbook);
    let image = read.worksheet(0).unwrap().images().next().expect("image");
    assert_eq!(
        image.object.meta.alt_text.as_deref(),
        Some("a tiny transparent pixel")
    );
}

/// Picture locked/print flags round-trip through the OBJ ftCmo grbit.
#[test]
fn xls_image_locked_printable_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let mut object = png("Pic").with_anchor(two_cell(0, 0, 2, 2));
    object.meta.locked = false;
    object.meta.printable = false;
    sheet.add_drawing(object);

    let read = round_trip(&workbook);
    let image = read.worksheet(0).unwrap().images().next().expect("image");
    assert!(!image.object.meta.locked);
    assert!(!image.object.meta.printable);

    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_drawing(png("Pic").with_anchor(two_cell(0, 0, 2, 2)));
    let read = round_trip(&workbook);
    let image = read.worksheet(0).unwrap().images().next().unwrap();
    assert!(image.object.meta.locked);
    assert!(image.object.meta.printable);
}
