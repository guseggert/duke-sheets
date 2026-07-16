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
    CellComment, CellMarker, ChildTransform, DrawingAnchor, DrawingKind, DrawingMeta,
    DrawingObject, EmbeddedImage, FormControl, FormControlKind, Group, GroupChild, GroupTransform,
    ImageFormat, Workbook,
};
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

fn checkbox(caption: &str, link: Option<&str>) -> FormControl {
    FormControl::new(FormControlKind::Checkbox {
        caption: caption.into(),
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
            DrawingKind::Shape(_) => "shape",
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
    sheet.add_drawing(png("Pic").with_anchor(two_cell(1, 1, 3, 3))).unwrap();

    let once = round_trip(&workbook);
    assert_eq!(kind_tags(&once), vec!["image"], "first round trip");

    let twice = round_trip(&once);
    assert_eq!(kind_tags(&twice), vec!["image"], "second round trip");
}

/// A form control's z-position among native drawing objects survives
/// a round trip via its a14 placeholder twin in the drawing part.
// features: Drawing z-order across kinds
#[test]
fn xlsx_control_z_order_between_images_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Middle", Some("$D$2")))
            .with_anchor(two_cell(1, 1, 3, 3)),
    ).unwrap();
    sheet.add_drawing(png("Above").with_anchor(two_cell(2, 2, 4, 4))).unwrap();

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["image", "control", "image"]);

    let sheet = read.worksheet(0).unwrap();
    let images: Vec<_> = sheet.images().collect();
    assert_eq!(images[0].object.unwrap().meta.name.as_deref(), Some("Below"));
    assert_eq!(images[1].object.unwrap().meta.name.as_deref(), Some("Above"));
    let control = sheet.form_controls().next().unwrap();
    assert_eq!(control.payload.caption_text().as_deref(), Some("Middle"));
    assert_eq!(
        control.payload.cell_link(),
        Some("$D$2"),
        "control state still comes from ctrlProps"
    );
    assert_eq!(control.object.unwrap().anchor, two_cell(1, 1, 3, 3));
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
    ).unwrap();
    sheet.add_drawing(DrawingObject::comment(
        4,
        4,
        CellComment::new("a", "between the controls"),
    )).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Second", None)).with_anchor(two_cell(6, 6, 8, 8)),
    ).unwrap();

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["control", "comment", "control"]);
    let sheet = read.worksheet(0).unwrap();
    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls[0].payload.caption_text().as_deref(), Some("First"));
    assert_eq!(
        controls[1].payload.caption_text().as_deref(),
        Some("Second")
    );
    assert_eq!(sheet.comment_at(4, 4).unwrap().plain_text(), "between the controls");
}

/// A shape group containing two pictures round-trips as a Group with
/// child transforms, instead of degrading to a raw blob.
// features: Grouped drawing objects
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
    ).unwrap();

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
    sheet.add_drawing(object).unwrap();

    let read = round_trip(&workbook);
    let sheet = read.worksheet(0).unwrap();
    let image = sheet.images().next().expect("image");
    assert!(!image.object.unwrap().meta.locked);
    assert!(!image.object.unwrap().meta.printable);

    // And the defaults still read back as true.
    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_drawing(png("Pic").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    let read = round_trip(&workbook);
    let image = read.worksheet(0).unwrap().images().next().unwrap();
    assert!(image.object.unwrap().meta.locked);
    assert!(image.object.unwrap().meta.printable);
}

/// A customized comment popup anchor survives a round trip instead of
/// being re-synthesized from the cell position.
// features: Plain-text comments with author; Comment positioning (anchor)
#[test]
fn xlsx_comment_anchor_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let custom = two_cell(8, 2, 12, 9);
    sheet.add_drawing(
        DrawingObject::comment(1, 1, CellComment::new("a", "moved popup"))
            .with_anchor(custom.clone()),
    ).unwrap();

    let read = round_trip(&workbook);
    let sheet = read.worksheet(0).unwrap();
    let comment = sheet.comments_drawn().next().expect("comment");
    assert_eq!((comment.row, comment.col), (1, 1));
    assert_eq!(comment.object.anchor, custom);
}

const RT_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/// A raw connector anchor whose cNvPr carries a hyperlink relationship
/// reference, plus the captured relationship it resolves through.
fn raw_link(col: u16, cnv_id: u32, rel_id: &str, url: &str) -> DrawingObject {
    let bytes = format!(
        r#"<xdr:twoCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{to}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="{cnv_id}" name="Connector {cnv_id}"><a:hlinkClick r:id="{rel_id}"/></xdr:cNvPr><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>"#,
        to = col + 1,
    );
    DrawingObject::raw(duke_sheets::RawDrawing {
        bytes: bytes.into_bytes(),
        rels: vec![duke_sheets::RawRel {
            id: rel_id.to_string(),
            rel_type: RT_HYPERLINK.to_string(),
            target: url.to_string(),
            external: true,
            part: None,
        }],
    })
    .with_anchor(two_cell(col, 1, col + 1, 2))
}

/// Raw fragments that reuse the same relationship id for different
/// targets get the colliding references remapped (in the .rels part
/// and inside the fragment bytes), while fragments sharing an id for
/// the same target keep sharing it.
#[test]
fn xlsx_conflicting_raw_rel_ids_are_remapped() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Pic").with_anchor(two_cell(8, 8, 9, 9))).unwrap();
    sheet.add_drawing(raw_link(0, 11, "rId1", "https://one.example/")).unwrap();
    sheet.add_drawing(raw_link(2, 12, "rId1", "https://two.example/")).unwrap();
    sheet.add_drawing(raw_link(4, 13, "rId1", "https://one.example/")).unwrap();

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["image", "raw", "raw", "raw"]);
    let sheet = read.worksheet(0).unwrap();
    let raws: Vec<&duke_sheets::RawDrawing> = sheet
        .drawings()
        .iter()
        .filter_map(|object| match &object.kind {
            DrawingKind::Raw(raw) => Some(raw),
            _ => None,
        })
        .collect();

    for (raw, url) in raws.iter().zip([
        "https://one.example/",
        "https://two.example/",
        "https://one.example/",
    ]) {
        assert_eq!(raw.rels.len(), 1, "one captured rel per fragment");
        assert_eq!(raw.rels[0].target, url, "fragment resolves its own target");
        let bytes = std::str::from_utf8(&raw.bytes).unwrap();
        assert!(
            bytes.contains(&format!("r:id=\"{}\"", raw.rels[0].id)),
            "fragment bytes reference the captured rel id"
        );
    }
    assert_eq!(raws[0].rels[0].id, "rId1", "first claimant keeps its id");
    assert_ne!(
        raws[1].rels[0].id, "rId1",
        "conflicting reuse is remapped to a fresh id"
    );
    assert_eq!(
        raws[2].rels[0].id, "rId1",
        "same-target reuse keeps sharing the id"
    );
}

/// A foreign file whose control twin carries a txBody for a
/// caption-less control kind (scrollbars, spinners, list boxes) must
/// not crash the reader; the stray text is ignored.
#[test]
fn xlsx_twin_text_on_captionless_control_is_ignored() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Scrollbar {
            value: 5,
            min: 0,
            max: 100,
            increment: 1,
            page: 10,
            horizontal: false,
            cell_link: None,
        }))
        .with_anchor(two_cell(1, 1, 2, 4)),
    ).unwrap();
    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut output).expect("write");
    let bytes = output.into_inner();

    // Splice a txBody into the scrollbar's twin sp, as a third-party
    // producer might.
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
    let patched = drawing.replace(
        "</xdr:sp>",
        "<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>stray</a:t></a:r></a:p></xdr:txBody></xdr:sp>",
    );
    assert_ne!(patched, drawing, "twin sp must be present to patch");

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
    let sheet = read.worksheet(0).unwrap();
    let controls = sheet.form_controls().collect::<Vec<_>>();
    assert_eq!(controls.len(), 1);
    assert!(
        matches!(
            controls[0].payload.kind,
            FormControlKind::Scrollbar { value: 5, .. }
        ),
        "scrollbar survives with its state"
    );
}

/// Raw (unmodeled) drawing anchors keep their position in the z-order
/// and expose a parsed anchor for placeholder rendering.
#[test]
fn xlsx_raw_anchor_keeps_order_and_parsed_anchor() {
    // Connector shapes are not modeled: build a file that contains
    // image / connector / image, then verify order and anchor.
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    sheet.add_drawing(png("Above").with_anchor(two_cell(4, 4, 6, 6))).unwrap();
    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut output).expect("write");
    let bytes = output.into_inner();

    // Splice a connector anchor between the two pics in drawing1.xml.
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
    let connector = r#"<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="99" name="Straight Connector 9"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm><a:off x="1828800" y="571500"/><a:ext cx="609600" cy="190500"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>"#;
    let insert_at = {
        let first_end = drawing.find("</xdr:twoCellAnchor>").expect("first anchor")
            + "</xdr:twoCellAnchor>".len();
        first_end
    };
    let mut patched = drawing.clone();
    patched.insert_str(insert_at, connector);

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

/// Rebuild a written package with a raw anchor spliced into
/// drawing1.xml, extra `<Relationship>` entries appended to
/// drawing1.xml.rels, and dummy target parts added.
fn splice_drawing_anchor(
    bytes: Vec<u8>,
    anchor: &str,
    extra_rels: &[(&str, &str, &str)],
    extra_parts: &[(&str, &str)],
) -> Vec<u8> {
    use std::io::Read;

    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let mut content = Vec::new();
        file.read_to_end(&mut content).unwrap();
        let options = zip::write::SimpleFileOptions::default();
        out.start_file(name.clone(), options).unwrap();
        if name == "xl/drawings/drawing1.xml" {
            let text = String::from_utf8(content).unwrap();
            let insert_at = text.find("</xdr:wsDr>").expect("wsDr end");
            let mut patched = text.clone();
            patched.insert_str(insert_at, anchor);
            out.write_all(patched.as_bytes()).unwrap();
        } else if name == "xl/drawings/_rels/drawing1.xml.rels" {
            let text = String::from_utf8(content).unwrap();
            let mut appended = String::new();
            for (id, rel_type, target) in extra_rels {
                appended.push_str(&format!(
                    r#"<Relationship Id="{id}" Type="{rel_type}" Target="{target}"/>"#
                ));
            }
            let patched = text.replace(
                "</Relationships>",
                &format!("{appended}</Relationships>"),
            );
            out.write_all(patched.as_bytes()).unwrap();
        } else {
            out.write_all(&content).unwrap();
        }
    }
    for (path, content) in extra_parts {
        let options = zip::write::SimpleFileOptions::default();
        out.start_file(path.to_string(), options).unwrap();
        out.write_all(content.as_bytes()).unwrap();
    }
    out.finish().unwrap().into_inner()
}

/// A raw SmartArt anchor references its four parts through
/// `dgm:relIds` attributes (r:dm/r:lo/r:qs/r:cs) rather than
/// r:id/r:embed/r:link. Every attribute whose value matches a rel id
/// in the drawing's .rels must be captured, with target parts, and
/// survive a round trip resolvable.
#[test]
fn xlsx_smartart_rel_ids_attributes_are_captured_and_round_trip() {
    const RT_DIAGRAM: [(&str, &str, &str); 4] = [
        (
            "rId2",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
            "../diagrams/data1.xml",
        ),
        (
            "rId3",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
            "../diagrams/layout1.xml",
        ),
        (
            "rId4",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
            "../diagrams/quickStyle1.xml",
        ),
        (
            "rId5",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
            "../diagrams/colors1.xml",
        ),
    ];
    const PARTS: [(&str, &str); 4] = [
        ("xl/diagrams/data1.xml", "<dataModelRoot/>"),
        ("xl/diagrams/layout1.xml", "<layoutDefRoot/>"),
        ("xl/diagrams/quickStyle1.xml", "<styleDefRoot/>"),
        ("xl/diagrams/colors1.xml", "<colorsDefRoot/>"),
    ];
    let anchor = r#"<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="7" name="Diagram 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"#;

    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_drawing(png("Pic").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut output).expect("write");
    let patched = splice_drawing_anchor(output.into_inner(), anchor, &RT_DIAGRAM, &PARTS);

    let assert_diagram_rels = |workbook: &Workbook, phase: &str| {
        let sheet_drawings = workbook.worksheet(0).unwrap();
        let raw = sheet_drawings
            .drawings()
            .iter()
            .find_map(|object| match &object.kind {
                DrawingKind::Raw(raw) => Some(raw),
                _ => None,
            })
            .unwrap_or_else(|| panic!("{phase}: raw diagram anchor"));
        assert_eq!(raw.rels.len(), 4, "{phase}: all four diagram rels captured");
        let bytes = std::str::from_utf8(&raw.bytes).unwrap();
        for ((_, rel_type, _), (_, part_body)) in RT_DIAGRAM.iter().zip(PARTS.iter()) {
            let rel = raw
                .rels
                .iter()
                .find(|rel| rel.rel_type == *rel_type)
                .unwrap_or_else(|| panic!("{phase}: rel {rel_type} captured"));
            assert!(!rel.external, "{phase}: diagram rels are internal");
            assert!(
                bytes.contains(&format!("\"{}\"", rel.id)),
                "{phase}: anchor bytes reference {} ({bytes})",
                rel.id
            );
            assert_eq!(
                rel.part.as_deref(),
                Some(part_body.as_bytes()),
                "{phase}: part bytes captured for {rel_type}"
            );
        }
    };

    let read = XlsxReader::read(Cursor::new(patched)).expect("read patched");
    assert_diagram_rels(&read, "first read");

    let again = round_trip(&read);
    assert_diagram_rels(&again, "round trip");
}

/// A sheet whose only form control lives inside a group must still
/// emit the legacy layers: the VML part with the control shape, the
/// worksheet `<controls>` block, and the ctrlProps part. Mirrors the
/// XLSB rels-alignment test.
// features: Grouped drawing objects
#[test]
fn xlsx_group_nested_control_emits_controls_block_and_vml() {
    use std::io::Read;

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::group(Group {
            transform: GroupTransform {
                cx_emu: 400_000,
                cy_emu: 400_000,
                child_cx_emu: 400_000,
                child_cy_emu: 400_000,
                ..GroupTransform::default()
            },
            children: vec![GroupChild {
                meta: DrawingMeta::default(),
                transform: ChildTransform {
                    cx_emu: 200_000,
                    cy_emu: 200_000,
                    ..ChildTransform::default()
                },
                kind: DrawingKind::FormControl(checkbox("In group", None)),
            }],
        })
        .with_anchor(two_cell(0, 0, 3, 3)),
    ).unwrap();

    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut output).expect("write");
    let bytes = output.into_inner();

    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).unwrap();
    let read_part = |archive: &mut zip::ZipArchive<Cursor<Vec<u8>>>, name: &str| -> String {
        let mut part = String::new();
        archive
            .by_name(name)
            .unwrap_or_else(|_| panic!("part {name} missing"))
            .read_to_string(&mut part)
            .unwrap();
        part
    };

    let worksheet = read_part(&mut archive, "xl/worksheets/sheet1.xml");
    assert!(
        worksheet.contains("<controls>"),
        "worksheet emits the <controls> block: {worksheet}"
    );
    assert!(
        worksheet.contains("<control "),
        "controls block carries the control entry"
    );

    // The ctrlProps part exists and holds the checkbox state.
    let ctrl_prop = read_part(&mut archive, "xl/ctrlProps/ctrlProp1.xml");
    assert!(
        ctrl_prop.contains("objectType=\"CheckBox\""),
        "ctrlProp part describes the checkbox: {ctrl_prop}"
    );

    // The VML part carries the control shape.
    let vml = read_part(&mut archive, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains("ObjectType=\"Checkbox\""),
        "control shape emitted in VML: {vml}"
    );

    // And the control survives the round trip inside the group.
    let read = XlsxReader::read(Cursor::new(bytes)).expect("read");
    let sheet = read.worksheet(0).unwrap();
    let controls: Vec<_> = sheet.form_controls().collect::<Vec<_>>();
    assert_eq!(controls.len(), 1, "group-nested control survives");
    let group = sheet
        .drawings()
        .iter()
        .find_map(|object| object.kind.as_group())
        .expect("group survives");
    assert!(
        group
            .children
            .iter()
            .any(|child| matches!(child.kind, DrawingKind::FormControl(_))),
        "control stays inside the group"
    );
}

/// A chartsheet drawing can carry raw anchors (e.g. a connector with
/// a hyperlink). Their relationships must be captured on read and
/// re-emitted on write, with the chart's rel id allocated around the
/// preserved ids, so both the chart and the raw reference resolve.
#[test]
fn xlsx_chartsheet_raw_anchor_rel_survives_and_chart_avoids_collision() {
    use duke_sheets::{Chart, ChartType, DataReference, DataSeries};
    use duke_sheets_core::{ChartSheet, SheetVisibility};
    use std::io::Read;

    const RT_HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    const URL: &str = "https://example.com/raw-anchor";
    let raw_anchor = r#"<xdr:absoluteAnchor><xdr:pos x="100" y="100"/><xdr:ext cx="600" cy="600"/><xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="42" name="Connector 42"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></xdr:cNvPr><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:absoluteAnchor>"#;

    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value_at(0, 0, 1.0)
        .unwrap();
    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    wb.add_chartsheet(ChartSheet {
        name: "Chart1".to_string(),
        chart,
        visibility: SheetVisibility::Visible,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    })
    .unwrap();
    let mut buf = Cursor::new(Vec::new());
    XlsxWriter::write(&wb, &mut buf).expect("write");

    // Rebuild the package so the chart sits at rId2 and a spliced raw
    // connector anchor holds a hyperlink at rId1.
    let mut archive = zip::ZipArchive::new(Cursor::new(buf.into_inner())).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let mut content = Vec::new();
        file.read_to_end(&mut content).unwrap();
        let options = zip::write::SimpleFileOptions::default();
        out.start_file(name.clone(), options).unwrap();
        if name == "xl/drawings/drawing1.xml" {
            let text = String::from_utf8(content).unwrap();
            assert!(text.contains(r#"r:id="rId1""#), "chart frame at rId1");
            let mut patched = text.replace(r#"r:id="rId1""#, r#"r:id="rId2""#);
            let insert_at = patched.find("</xdr:wsDr>").expect("wsDr end");
            patched.insert_str(insert_at, raw_anchor);
            out.write_all(patched.as_bytes()).unwrap();
        } else if name == "xl/drawings/_rels/drawing1.xml.rels" {
            let text = String::from_utf8(content).unwrap();
            let patched = text.replace("Id=\"rId1\"", "Id=\"rId2\"").replace(
                "</Relationships>",
                &format!(
                    r#"<Relationship Id="rId1" Type="{RT_HYPERLINK}" Target="{URL}" TargetMode="External"/></Relationships>"#
                ),
            );
            out.write_all(patched.as_bytes()).unwrap();
        } else {
            out.write_all(&content).unwrap();
        }
    }
    let patched = out.finish().unwrap().into_inner();

    let read = XlsxReader::read(Cursor::new(patched)).expect("read patched");
    let cs = read.chartsheet(0).unwrap();
    assert_eq!(cs.chart.chart_type, ChartType::ColumnClustered);
    assert!(
        cs.raw_drawing_objects
            .iter()
            .any(|entry| std::str::from_utf8(entry).unwrap().contains("hlinkClick")),
        "raw connector anchor preserved"
    );

    // Round trip; both references must resolve in the output package.
    let mut out2 = Cursor::new(Vec::new());
    XlsxWriter::write(&read, &mut out2).expect("rewrite");
    let bytes2 = out2.into_inner();

    let mut archive = zip::ZipArchive::new(Cursor::new(bytes2.clone())).unwrap();
    let read_part = |archive: &mut zip::ZipArchive<Cursor<Vec<u8>>>, name: &str| -> String {
        let mut content = String::new();
        archive
            .by_name(name)
            .unwrap_or_else(|_| panic!("part {name} missing"))
            .read_to_string(&mut content)
            .unwrap();
        content
    };
    let names: Vec<String> = archive.file_names().map(str::to_string).collect();
    let drawing_path = names
        .iter()
        .filter(|name| name.starts_with("xl/drawings/drawing") && name.ends_with(".xml"))
        .find(|name| read_part(&mut archive, name).contains("hlinkClick"))
        .expect("chartsheet drawing part with the raw anchor")
        .clone();
    let drawing = read_part(&mut archive, &drawing_path);
    let attr_id = |text: &str, marker: &str| -> String {
        let at = text
            .find(marker)
            .unwrap_or_else(|| panic!("{marker} in {text}"));
        let rest = &text[at..];
        let start = rest.find("r:id=\"").expect("r:id attr") + 6;
        rest[start..].split('"').next().unwrap().to_string()
    };
    let hlink_id = attr_id(&drawing, "hlinkClick");
    let chart_id = attr_id(&drawing, "<c:chart");
    assert_ne!(
        hlink_id, chart_id,
        "chart rel id must not collide with the raw anchor's reference"
    );

    let rels_path = drawing_path.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
    let rels = read_part(&mut archive, &rels_path);
    let rel_entry = |rid: &str| -> String {
        let marker = format!("Id=\"{rid}\"");
        rels.split("<Relationship ")
            .find(|chunk| chunk.contains(&marker))
            .unwrap_or_else(|| panic!("relationship {rid} missing in {rels}"))
            .to_string()
    };
    let hlink_rel = rel_entry(&hlink_id);
    assert!(
        hlink_rel.contains(&format!("Target=\"{URL}\"")),
        "hyperlink target preserved: {hlink_rel}"
    );
    assert!(
        hlink_rel.contains(RT_HYPERLINK),
        "hyperlink type preserved: {hlink_rel}"
    );
    assert!(
        rel_entry(&chart_id).contains("/chart"),
        "chart relationship resolves"
    );

    // And the rewritten package still reads with both intact.
    let read2 = XlsxReader::read(Cursor::new(bytes2)).expect("read rewrite");
    let cs2 = read2.chartsheet(0).unwrap();
    assert_eq!(cs2.chart.chart_type, ChartType::ColumnClustered);
    assert!(
        cs2.raw_drawing_objects
            .iter()
            .any(|entry| std::str::from_utf8(entry).unwrap().contains("hlinkClick")),
        "raw connector anchor survives the second read"
    );
}

/// A malformed drawing part must not fail the whole workbook read:
/// the drawing is skipped (matching the XLSB reader) and the cell
/// data stays readable.
#[test]
fn xlsx_malformed_drawing_part_is_skipped() {
    use std::io::Read;

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value_at(0, 0, "kept").unwrap();
    sheet.add_drawing(png("Pic").with_anchor(two_cell(1, 1, 3, 3))).unwrap();
    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut output).expect("write");

    // Corrupt drawing1.xml into non-XML garbage.
    let mut archive = zip::ZipArchive::new(Cursor::new(output.into_inner())).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let options = zip::write::SimpleFileOptions::default();
        out.start_file(name.clone(), options).unwrap();
        if name == "xl/drawings/drawing1.xml" {
            out.write_all(b"<xdr:wsDr><broken").unwrap();
        } else {
            let mut content = Vec::new();
            file.read_to_end(&mut content).unwrap();
            out.write_all(&content).unwrap();
        }
    }
    let corrupted = out.finish().unwrap().into_inner();

    let read = XlsxReader::read(Cursor::new(corrupted)).expect("read succeeds");
    let sheet = read.worksheet(0).unwrap();
    assert_eq!(
        sheet.get_value_at(0, 0),
        duke_sheets::CellValue::string("kept")
    );
    assert!(
        sheet.drawings().is_empty(),
        "malformed drawing part yields zero drawings"
    );
}

/// Chart anchors keep their variant across a round trip: a
/// OneCell-anchored chart must not collapse to an all-zero TwoCell
/// anchor, and editAs on a TwoCell chart anchor survives.
#[test]
fn xlsx_chart_one_cell_and_edit_as_anchors_round_trip() {
    use duke_sheets::{Chart, ChartType, DataReference, DataSeries, EditAs};

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    for (i, v) in [1.0, 4.0, 2.0].iter().enumerate() {
        sheet.set_cell_value_at(i as u32, 0, *v).unwrap();
    }

    let mut one_cell_chart = Chart::new(ChartType::ColumnClustered);
    one_cell_chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
    let one_cell = DrawingAnchor::OneCell {
        from: CellMarker {
            col: 1,
            col_offset_emu: 95_250,
            row: 2,
            row_offset_emu: 47_625,
        },
        width_emu: 5_000_000,
        height_emu: 3_000_000,
    };
    sheet.add_chart(one_cell_chart, one_cell.clone()).unwrap();

    let mut pinned_chart = Chart::new(ChartType::ColumnClustered);
    pinned_chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
    let pinned = DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 4,
            col_offset_emu: 0,
            row: 4,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: 10,
            col_offset_emu: 0,
            row: 18,
            row_offset_emu: 0,
        },
        edit_as: Some(EditAs::Absolute),
    };
    sheet.add_chart(pinned_chart, pinned.clone()).unwrap();

    let read = round_trip(&workbook);
    let charts: Vec<_> = read.worksheet(0).unwrap().charts().collect();
    assert_eq!(charts.len(), 2);
    assert_eq!(
        charts[0].object.unwrap().anchor, one_cell,
        "OneCell chart anchor survives with its extent"
    );
    assert_eq!(
        charts[1].object.unwrap().anchor, pinned,
        "TwoCell chart anchor keeps editAs"
    );
}

/// An anonymous (empty-author) comment keeps its own author slot
/// instead of being attributed to the first named author.
#[test]
fn xlsx_empty_author_comment_keeps_attribution() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_comment_at(0, 0, CellComment::new("Bob", "named")).unwrap();
    sheet.set_comment_at(1, 0, CellComment::new("", "anonymous")).unwrap();

    let read = round_trip(&workbook);
    let sheet = read.worksheet(0).unwrap();
    assert_eq!(sheet.comment_at(0, 0).unwrap().author, "Bob");
    assert_eq!(
        sheet.comment_at(1, 0).unwrap().author,
        "",
        "anonymous comment must not inherit another author"
    );
}
