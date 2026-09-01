//! Round-trip tests for BIFF8 shape groups (`SpgrContainer`).
//!
//! The flat acceptance coverage for groups lives in
//! `duke-sheets/tests/xls_drawing_order.rs`; this file exercises the
//! deeper structures: nested groups, grouped captioned controls (a
//! grouped shape's ClientTextbox/TXO interleave), radio chains that
//! span group membership, and group metadata flags.

use std::io::Cursor;

use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
use duke_sheets_core::{
    CheckState, ChildTransform, DrawingKind, DrawingMeta, DrawingObject, FormControl,
    FormControlKind, Group, GroupChild, GroupTransform, Workbook,
};
use duke_sheets_xls::{XlsReader, XlsWriter};

const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82,
];

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn anchor(from_col: u16, from_row: u32, to_col: u16, to_row: u32) -> DrawingAnchor {
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

fn png_image() -> EmbeddedImage {
    EmbeddedImage {
        format: ImageFormat::Png,
        media_path: String::new(),
        svg_media_path: None,
        width_emu: 190_500,
        height_emu: 190_500,
        rotation: None,
        flip_h: false,
        flip_v: false,
        data: TEST_PNG_1X1.to_vec(),
        svg_data: None,
    }
}

fn child(name: &str, x: i64, y: i64, kind: DrawingKind) -> GroupChild {
    GroupChild {
        meta: DrawingMeta {
            name: Some(name.to_string()),
            ..Default::default()
        },
        transform: ChildTransform {
            x_emu: x,
            y_emu: y,
            cx_emu: 190_500,
            cy_emu: 190_500,
            rotation: 0,
            flip_h: false,
            flip_v: false,
        },
        kind,
    }
}

fn group_transform(cx: i64, cy: i64) -> GroupTransform {
    GroupTransform {
        x_emu: 0,
        y_emu: 0,
        cx_emu: cx,
        cy_emu: cy,
        child_x_emu: 0,
        child_y_emu: 0,
        child_cx_emu: cx,
        child_cy_emu: cy,
        rotation: 0,
        flip_h: false,
        flip_v: false,
    }
}

fn checkbox(caption: &str) -> FormControlKind {
    FormControlKind::Checkbox {
        caption: caption.into(),
        state: CheckState::Checked,
        cell_link: None,
        no_3d: false,
    }
}

#[test]
fn grouped_captioned_control_round_trips() {
    // A grouped control routes its caption through the same
    // ClientTextbox-after-OBJ TXO interleave as a top-level one,
    // inside the SpgrContainer.
    let mut wb = Workbook::new();
    let group = Group {
        transform: group_transform(1_219_200, 190_500),
        children: vec![
            child(
                "Grouped check",
                0,
                0,
                DrawingKind::FormControl(FormControl::new(checkbox("In group"))),
            ),
            child("Pic", 609_600, 0, DrawingKind::Image(png_image())),
        ],
    };
    wb.worksheet_mut(0)
        .unwrap()
        .add_drawing(DrawingObject::group(group).with_anchor(anchor(1, 1, 4, 2))).unwrap();

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let group = sheet.drawings()[0].kind.as_group().expect("group");
    assert_eq!(group.children.len(), 2);
    let control = group.children[0]
        .kind
        .as_form_control()
        .expect("grouped control");
    assert_eq!(control.caption_text().as_deref(), Some("In group"));
    assert_eq!(
        group.children[0].meta.name.as_deref(),
        Some("Grouped check"),
        "grouped control wzName survives"
    );
    assert!(matches!(group.children[1].kind, DrawingKind::Image(_)));
}

#[test]
fn nested_group_round_trips() {
    let inner = Group {
        transform: group_transform(190_500, 190_500),
        children: vec![child("Deep pic", 0, 0, DrawingKind::Image(png_image()))],
    };
    let outer = Group {
        transform: group_transform(1_219_200, 190_500),
        children: vec![
            child(
                "Inner group",
                609_600,
                0,
                DrawingKind::Group(Box::new(inner)),
            ),
            child("Outer pic", 0, 0, DrawingKind::Image(png_image())),
        ],
    };
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .add_drawing(DrawingObject::group(outer).with_anchor(anchor(1, 1, 4, 2))).unwrap();

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let outer = sheet.drawings()[0].kind.as_group().expect("outer group");
    assert_eq!(outer.children.len(), 2);
    let nested = outer.children[0].kind.as_group().expect("nested group");
    assert_eq!(outer.children[0].transform.x_emu, 609_600);
    assert_eq!(nested.children.len(), 1);
    assert_eq!(nested.children[0].meta.name.as_deref(), Some("Deep pic"));
    assert!(matches!(nested.children[0].kind, DrawingKind::Image(_)));
    assert!(
        matches!(outer.children[1].kind, DrawingKind::Image(_)),
        "sibling after the nested group keeps its pairing"
    );
}

// features: Grouped drawing objects
#[test]
fn group_and_grouped_picture_rotation_round_trip_in_model_units() {
    // FOPT 0x0004 is FixedPoint 16.16 degrees on the wire (MS-ODRAW
    // 2.3.18.5); the model stays in 60,000ths of a degree. 90 degrees
    // = 5,400,000 model units must survive for the group transform,
    // a grouped picture's child transform, and a grouped shape.
    let mut rotated_shape = duke_sheets_core::Shape::preset("rect");
    rotated_shape.rotation = 2_700_000; // 45 degrees
    let mut group = Group {
        transform: group_transform(1_219_200, 190_500),
        children: vec![
            child("Pic", 0, 0, DrawingKind::Image(png_image())),
            child(
                "Rect",
                609_600,
                0,
                DrawingKind::Shape(Box::new(rotated_shape)),
            ),
        ],
    };
    group.transform.rotation = 5_400_000;
    group.children[0].transform.rotation = 5_400_000;
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .add_drawing(DrawingObject::group(group).with_anchor(anchor(1, 1, 4, 2))).unwrap();

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let group = sheet.drawings()[0].kind.as_group().expect("group");
    assert_eq!(group.transform.rotation, 5_400_000, "group rotation");
    assert_eq!(
        group.children[0].transform.rotation, 5_400_000,
        "grouped picture child rotation"
    );
    let shape = match &group.children[1].kind {
        DrawingKind::Shape(shape) => shape,
        other => panic!("expected grouped shape, got {other:?}"),
    };
    assert_eq!(shape.rotation, 2_700_000, "grouped shape rotation");
}

#[test]
fn group_locked_printable_flags_round_trip() {
    let group = Group {
        transform: group_transform(609_600, 190_500),
        children: vec![child("P", 0, 0, DrawingKind::Image(png_image()))],
    };
    let mut wb = Workbook::new();
    let mut object = DrawingObject::group(group).with_anchor(anchor(1, 1, 2, 2));
    object.meta.locked = false;
    object.meta.printable = false;
    wb.worksheet_mut(0).unwrap().add_drawing(object).unwrap();

    let parsed = write_then_read(&wb);
    let object = &parsed.worksheet(0).unwrap().drawings()[0];
    assert!(object.kind.as_group().is_some());
    assert!(!object.meta.locked, "group fLocked survives the OBJ grbit");
    assert!(
        !object.meta.printable,
        "group fPrint survives the OBJ grbit"
    );
}

#[test]
fn radio_chain_spans_grouped_and_top_level_radios() {
    // Radios inside a shape group join the sheet's FtRboData chains;
    // grouping is computed from resolved on-sheet rectangles, so a
    // grouped radio inside a group box chains with it.
    let radio = |caption: &str| FormControlKind::OptionButton {
        caption: caption.into(),
            state: CheckState::Unchecked,
            cell_link: None,
            first_in_group: false,
            no_3d: false,
    };
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    // A shape group spanning A1:C4 holding one radio.
    let group = Group {
        transform: group_transform(1_828_800, 762_000),
        children: vec![child(
            "R1",
            0,
            0,
            DrawingKind::FormControl(FormControl::new(radio("Grouped"))),
        )],
    };
    ws.add_drawing(DrawingObject::group(group).with_anchor(anchor(0, 0, 3, 4))).unwrap();
    // A top-level radio elsewhere on the sheet.
    ws.add_drawing(
        DrawingObject::form_control(FormControl::new(radio("Loose")))
            .with_anchor(anchor(6, 6, 8, 8)),
    ).unwrap();

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let group = sheet.drawings()[0].kind.as_group().expect("group");
    match &group.children[0].kind {
        DrawingKind::FormControl(control) => match &control.kind {
            FormControlKind::OptionButton { first_in_group, .. } => {
                assert!(
                    first_in_group,
                    "first radio of the sheet-level chain carries fFirstBtn"
                );
            }
            other => panic!("expected OptionButton, got {other:?}"),
        },
        other => panic!("expected grouped control, got {other:?}"),
    }
}

#[test]
fn shape_text_default_alignment_round_trips_to_none() {
    // Shape text with no explicit alignment is written with the
    // Left/Top TXO defaults; the reader must strip those back to
    // None the way the control caption path does.
    let mut shape = duke_sheets_core::Shape::preset("rect");
    shape.text = Some("hello".into());
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .add_shape(shape, anchor(1, 1, 3, 3)).unwrap();

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let drawn = sheet.shapes().next().expect("shape survives");
    let text = drawn.payload.text.as_ref().expect("shape text");
    assert_eq!(text.plain_text(), "hello");
    assert_eq!(
        text.horizontal_alignment, None,
        "default horizontal alignment must strip to None"
    );
    assert_eq!(
        text.vertical_alignment, None,
        "default vertical alignment must strip to None"
    );
}

#[test]
fn unmodelable_group_children_are_dropped_not_the_group() {
    // Comments cannot live inside an XLS shape group; the writer
    // drops the comment child and keeps the rest of the group.
    let mut children = vec![child("Keep", 0, 0, DrawingKind::Image(png_image()))];
    children.push(GroupChild {
        meta: DrawingMeta::default(),
        transform: ChildTransform::default(),
        kind: DrawingKind::Comment {
            row: 9,
            col: 9,
            comment: duke_sheets_core::CellComment::new("a", "grouped note"),
        },
    });
    let group = Group {
        transform: group_transform(609_600, 190_500),
        children,
    };
    let mut wb = Workbook::new();
    // Unchecked: a comment group child fails validation on purpose;
    // the writer must still drop it for files read permissively.
    wb.worksheet_mut(0)
        .unwrap()
        .drawings_mut()
        .push(DrawingObject::group(group).with_anchor(anchor(1, 1, 2, 2)));

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let group = sheet.drawings()[0].kind.as_group().expect("group");
    assert_eq!(group.children.len(), 1, "comment child is dropped");
    assert_eq!(group.children[0].meta.name.as_deref(), Some("Keep"));
}

/// LibreOffice envelope check: write an XLS whose drawing stream
/// interleaves a group (with a grouped picture + grouped control), a
/// comment, and a top-level control, open via URP, and read the
/// anchor cell back. Catches structural malformations in the
/// restructured SpgrContainer / OBJ / TXO interleave that make LO's
/// loader refuse the file.
#[test]
fn lo_can_open_xls_with_shape_group_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");
    let group = Group {
        transform: group_transform(1_219_200, 190_500),
        children: vec![
            child("G pic", 0, 0, DrawingKind::Image(png_image())),
            child(
                "G check",
                609_600,
                0,
                DrawingKind::FormControl(FormControl::new(checkbox("Grouped"))),
            ),
        ],
    };
    ws.add_drawing(
        DrawingObject::group(group)
            .with_anchor(anchor(1, 1, 4, 2))
            .with_name("Group 1"),
    ).unwrap();
    ws.set_comment_at(3, 3, duke_sheets_core::CellComment::new("a", "note"))
        .expect("set comment");
    ws.add_drawing(
        DrawingObject::form_control(FormControl::new(checkbox("Top level")))
            .with_anchor(anchor(5, 5, 7, 7)),
    ).unwrap();

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_shape_group_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<f64, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb_in = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        wb_in
            .get_cell_value("A1")
            .await
            .map_err(|e| format!("A1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    let a1 = outcome.expect("LO must open our XLS with a shape group without error");
    assert!(
        (a1 - 42.0).abs() < 1e-9,
        "A1 must round-trip; got {a1} (expected 42)"
    );
}
