//! XLSX drawing z-order, grouping, and shared-property fidelity.
//!
//! The drawings list order is z-order (back to front). In XLSX the
//! native drawing part's document order carries it; form controls
//! participate via their a14 placeholder twins, and comments (legacy
//! VML only) keep their order relative to controls. Cross-layer order
//! between comments and native objects normalizes on round-trip; all
//! other relative order must survive exactly.

use std::io::Cursor;

use duke_sheets::{
    CellComment, CellMarker, ChildTransform, DrawingAnchor, DrawingKind, DrawingObject,
    EmbeddedImage, FormControl, FormControlKind, Group, GroupChild, GroupTransform, ImageFormat,
    Workbook,
};
use duke_sheets_xlsx::{XlsxReader, XlsxWriter};

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

fn checkbox(caption: &str, link: Option<&str>) -> FormControl {
    FormControl::new(FormControlKind::Checkbox {
        caption: caption.to_string(),
        state: duke_sheets::CheckState::Checked,
        cell_link: link.map(str::to_string),
        no_3d: false,
    })
}

fn round_trip(workbook: &Workbook) -> Workbook {
    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(workbook, &mut output).expect("write");
    XlsxReader::read(Cursor::new(output.into_inner())).expect("read")
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

/// Write -> read -> write -> read must not duplicate images. The old
/// reader stored every pic anchor twice (parsed image + raw anchor),
/// doubling images on each cycle.
#[test]
fn xlsx_image_double_round_trip_does_not_duplicate() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Pic").with_anchor(two_cell(1, 1, 3, 3)));

    let once = round_trip(&workbook);
    assert_eq!(kind_tags(&once), vec!["image"], "first round trip");

    let twice = round_trip(&once);
    assert_eq!(kind_tags(&twice), vec!["image"], "second round trip");
}

/// A form control's z-position among native drawing objects survives
/// a round trip via its a14 placeholder twin in the drawing part.
#[test]
fn xlsx_control_z_order_between_images_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2)));
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Middle", Some("$D$2")))
            .with_anchor(two_cell(1, 1, 3, 3)),
    );
    sheet.add_drawing(png("Above").with_anchor(two_cell(2, 2, 4, 4)));

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["image", "control", "image"]);

    let sheet = read.worksheet(0).unwrap();
    let images: Vec<_> = sheet.images().collect();
    assert_eq!(images[0].object.meta.name.as_deref(), Some("Below"));
    assert_eq!(images[1].object.meta.name.as_deref(), Some("Above"));
    let control = sheet.form_controls().next().unwrap();
    assert_eq!(control.payload.caption(), Some("Middle"));
    assert_eq!(
        control.payload.cell_link(),
        Some("$D$2"),
        "control state still comes from ctrlProps"
    );
    assert_eq!(control.object.anchor, two_cell(1, 1, 3, 3));
}

/// Comments keep their order relative to form controls (the shared
/// legacy VML layer), and comment-vs-native order normalizes without
/// dropping anything.
#[test]
fn xlsx_comment_control_relative_order_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("First", None)).with_anchor(two_cell(0, 0, 2, 2)),
    );
    sheet.add_drawing(DrawingObject::comment(
        4,
        4,
        CellComment::new("a", "between the controls"),
    ));
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Second", None)).with_anchor(two_cell(6, 6, 8, 8)),
    );

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["control", "comment", "control"]);
    let sheet = read.worksheet(0).unwrap();
    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls[0].payload.caption(), Some("First"));
    assert_eq!(controls[1].payload.caption(), Some("Second"));
    assert_eq!(sheet.comment_at(4, 4).unwrap().text, "between the controls");
}

/// A shape group containing two pictures round-trips as a Group with
/// child transforms, instead of degrading to a raw blob.
#[test]
fn xlsx_group_of_images_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();

    let child = |name: &str, x: i64| GroupChild {
        meta: duke_sheets::DrawingMeta {
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

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["group"]);
    let sheet = read.worksheet(0).unwrap();
    let object = &sheet.drawings()[0];
    assert_eq!(object.meta.name.as_deref(), Some("Group 1"));
    assert_eq!(object.anchor, two_cell(1, 1, 4, 2));
    let group = object.kind.as_group().expect("group payload");
    assert_eq!(group.transform.child_cx_emu, 1_219_200);
    assert_eq!(group.children.len(), 2);
    assert_eq!(group.children[0].meta.name.as_deref(), Some("Left"));
    assert_eq!(group.children[1].meta.name.as_deref(), Some("Right"));
    assert_eq!(group.children[1].transform.x_emu, 609_600);
    assert!(
        matches!(group.children[0].kind, DrawingKind::Image(_)),
        "children are modeled images"
    );
}

/// Image locked/printable flags round-trip through
/// xdr:clientData@fLocksWithSheet/fPrintsWithSheet.
#[test]
fn xlsx_image_locked_printable_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let mut object = png("Pic").with_anchor(two_cell(0, 0, 2, 2));
    object.meta.locked = false;
    object.meta.printable = false;
    sheet.add_drawing(object);

    let read = round_trip(&workbook);
    let sheet = read.worksheet(0).unwrap();
    let image = sheet.images().next().expect("image");
    assert!(!image.object.meta.locked);
    assert!(!image.object.meta.printable);

    // And the defaults still read back as true.
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

/// A customized comment popup anchor survives a round trip instead of
/// being re-synthesized from the cell position.
#[test]
fn xlsx_comment_anchor_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let custom = two_cell(8, 2, 12, 9);
    sheet.add_drawing(
        DrawingObject::comment(1, 1, CellComment::new("a", "moved popup")).with_anchor(custom.clone()),
    );

    let read = round_trip(&workbook);
    let sheet = read.worksheet(0).unwrap();
    let comment = sheet.comments_drawn().next().expect("comment");
    assert_eq!((comment.row, comment.col), (1, 1));
    assert_eq!(comment.object.anchor, custom);
}

/// Raw (unmodeled) drawing anchors keep their position in the z-order
/// and expose a parsed anchor for placeholder rendering.
#[test]
fn xlsx_raw_anchor_keeps_order_and_parsed_anchor() {
    // A plain textbox shape is not modeled yet: build a file that
    // contains image / shape / image, then verify order and anchor.
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2)));
    sheet.add_drawing(png("Above").with_anchor(two_cell(4, 4, 6, 6)));
    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut output).expect("write");
    let bytes = output.into_inner();

    // Splice a textbox anchor between the two pics in drawing1.xml.
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut drawing = String::new();
    {
        use std::io::Read;
        archive
            .by_name("xl/drawings/drawing1.xml")
            .unwrap()
            .read_to_string(&mut drawing)
            .unwrap();
    }
    let textbox = r#"<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="99" name="TextBox 9"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="1828800" y="571500"/><a:ext cx="609600" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>hello</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"#;
    let insert_at = {
        let first_end = drawing.find("</xdr:twoCellAnchor>").expect("first anchor") +
            "</xdr:twoCellAnchor>".len();
        first_end
    };
    let mut patched = drawing.clone();
    patched.insert_str(insert_at, textbox);

    // Rebuild the archive with the patched drawing part.
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        use std::io::Read;
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let options = zip::write::SimpleFileOptions::default();
        out.start_file(name.clone(), options).unwrap();
        if name == "xl/drawings/drawing1.xml" {
            out.write_all(patched.as_bytes()).unwrap();
        } else {
            let mut content = Vec::new();
            file.read_to_end(&mut content).unwrap();
            out.write_all(&content).unwrap();
        }
    }
    let patched_bytes = out.finish().unwrap().into_inner();

    let read = XlsxReader::read(Cursor::new(patched_bytes)).expect("read patched");
    assert_eq!(kind_tags(&read), vec!["image", "raw", "image"]);
    let sheet = read.worksheet(0).unwrap();
    let raw_object = &sheet.drawings()[1];
    assert_eq!(
        raw_object.anchor,
        two_cell(3, 3, 4, 4),
        "raw anchor parsed for placeholder rendering"
    );

    // And it survives a rewrite in place.
    let again = round_trip(&read);
    assert_eq!(kind_tags(&again), vec!["image", "raw", "image"]);
}

use std::io::Write;
