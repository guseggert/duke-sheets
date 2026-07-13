//! XLSB drawing z-order, grouping, chart persistence, and comment
//! record fidelity.
//!
//! XLSB shares the OOXML drawing part with XLSX, so z-order rides the
//! part's document order and controls participate via placeholder
//! twins (the `com14:compatSp` graphicFrame flavor, unlike XLSX's
//! `a14:compatExt` sp flavor). Historically the XLSB writer replayed
//! an opaque bundle captured at read time and never wrote model
//! drawings at all, silently discarding chart and image edits; and
//! the comments part used off-spec record ids (0x0278-based instead
//! of 0x0274-based per MS-XLSB 2.4.33) that Excel refuses.

use std::io::Cursor;

use duke_sheets::{
    CellComment, CellMarker, Chart, ChartType, ChildTransform, DataReference, DataSeries,
    DrawingAnchor, DrawingKind, DrawingMeta, DrawingObject, EmbeddedImage, FormControl,
    FormControlKind, Group, GroupChild, GroupTransform, ImageFormat, Workbook,
};
use duke_sheets_xlsb::{XlsbReader, XlsbWriter};

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

fn write_bytes(workbook: &Workbook) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    XlsbWriter::write(workbook, &mut cursor).expect("write");
    cursor.into_inner()
}

fn round_trip(workbook: &Workbook) -> Workbook {
    XlsbReader::read(Cursor::new(write_bytes(workbook))).expect("read")
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
fn xlsb_conflicting_raw_rel_ids_are_remapped() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Pic").with_anchor(two_cell(8, 8, 9, 9)));
    sheet.add_drawing(raw_link(0, 11, "rId1", "https://one.example/"));
    sheet.add_drawing(raw_link(2, 12, "rId1", "https://two.example/"));
    sheet.add_drawing(raw_link(4, 13, "rId1", "https://one.example/"));

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

/// A model-added image must actually be written to XLSB (the old
/// writer only replayed bundles captured at read time, so
/// model-authored drawings vanished and the dangling BrtDrawing
/// pointer made Excel refuse the file).
// features: Image parsing (PNG/JPEG); Image positioning (two-cell anchor)
#[test]
fn xlsb_model_image_round_trips() {
    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_drawing(png("Pic").with_anchor(two_cell(1, 1, 3, 3)));

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["image"]);
    let image = read.worksheet(0).unwrap().images().next().expect("image");
    assert_eq!(image.payload.data, PNG_1PX);
    assert_eq!(image.object.meta.name.as_deref(), Some("Pic"));
    assert_eq!(image.object.anchor, two_cell(1, 1, 3, 3));
}

// features: Image positioning (one-cell anchor); Image positioning (absolute anchor); Image editAs (move/size with cells)
#[test]
fn xlsb_one_cell_and_absolute_image_anchors_round_trip() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        png("OneCell").with_anchor(DrawingAnchor::OneCell {
            from: CellMarker {
                col: 1,
                row: 2,
                col_offset_emu: 95_250,
                row_offset_emu: 47_625,
            },
            width_emu: 1_200_000,
            height_emu: 700_000,
        }),
    );
    sheet.add_drawing(
        png("Absolute").with_anchor(DrawingAnchor::Absolute {
            x_emu: 2_000_000,
            y_emu: 1_000_000,
            width_emu: 900_000,
            height_emu: 500_000,
        }),
    );

    let read = round_trip(&workbook);
    let images: Vec<_> = read.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 2);
    assert_eq!(
        images[0].object.anchor,
        DrawingAnchor::OneCell {
            from: CellMarker {
                col: 1,
                row: 2,
                col_offset_emu: 95_250,
                row_offset_emu: 47_625,
            },
            width_emu: 1_200_000,
            height_emu: 700_000,
        }
    );
    assert_eq!(
        images[1].object.anchor,
        DrawingAnchor::Absolute {
            x_emu: 2_000_000,
            y_emu: 1_000_000,
            width_emu: 900_000,
            height_emu: 500_000,
        }
    );
}

/// A model-added chart survives an XLSB round trip, and chart edits
/// made after reading persist through the next write (the bundle
/// replay silently discarded them).
#[test]
fn xlsb_chart_edits_persist() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    for (i, v) in [1.0, 4.0, 2.0].iter().enumerate() {
        sheet.set_cell_value_at(i as u32, 0, *v).unwrap();
    }
    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.title = Some("Original".to_string());
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
    sheet.add_chart(chart, two_cell(2, 2, 8, 12));

    let mut read = round_trip(&workbook);
    {
        let sheet = read.worksheet_mut(0).unwrap();
        let chart = sheet.charts_mut().next().expect("chart survived");
        chart.title = Some("Edited".to_string());
    }
    let again = round_trip(&read);
    let chart = again.worksheet(0).unwrap().charts().next().expect("chart");
    assert_eq!(chart.payload.title.as_deref(), Some("Edited"));
    assert_eq!(chart.object.anchor, two_cell(2, 2, 8, 12));
}

/// Control z-position among native objects survives via the
/// com14:compatSp placeholder twin.
// features: Drawing z-order across kinds
#[test]
fn xlsb_control_z_order_between_images_round_trips() {
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
    assert_eq!(control.payload.caption_text().as_deref(), Some("Middle"));
    assert_eq!(control.object.anchor, two_cell(1, 1, 3, 3));
}

/// Control names ride the twin's cNvPr and survive the round trip
/// (previously dropped: XLSB has no ctrlProps part and the VML shape
/// carries no name).
#[test]
fn xlsb_control_name_round_trips() {
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(
        DrawingObject::form_control(checkbox("Named"))
            .with_anchor(two_cell(1, 1, 3, 3))
            .with_name("Check Box 9"),
    );

    let read = round_trip(&workbook);
    let control = read
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("control");
    assert_eq!(control.object.meta.name.as_deref(), Some("Check Box 9"));
}

/// Comments keep their order relative to controls via the shared
/// legacy VML sequence.
#[test]
fn xlsb_comment_control_relative_order_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("First")).with_anchor(two_cell(0, 0, 2, 2)),
    );
    sheet.add_drawing(DrawingObject::comment(
        4,
        4,
        CellComment::new("a", "between the controls"),
    ));
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Second")).with_anchor(two_cell(6, 6, 8, 8)),
    );

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["control", "comment", "control"]);
    assert_eq!(
        read.worksheet(0).unwrap().comment_at(4, 4).unwrap().text,
        "between the controls"
    );
}

/// Groups round-trip as a modeled tree through the shared drawing
/// part XML.
// features: Grouped drawing objects
#[test]
fn xlsb_group_of_images_round_trips() {
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
    sheet.add_drawing(
        DrawingObject::group(Group {
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
        })
        .with_anchor(two_cell(1, 1, 4, 2))
        .with_name("Group 1"),
    );

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["group"]);
    let object = &read.worksheet(0).unwrap().drawings()[0];
    let group = object.kind.as_group().expect("group");
    assert_eq!(group.children.len(), 2);
    assert_eq!(group.children[0].meta.name.as_deref(), Some("Left"));
    assert_eq!(group.children[1].transform.x_emu, 609_600);
}

/// A customized comment popup anchor survives the round trip.
// features: Comment positioning (anchor)
#[test]
fn xlsb_comment_anchor_round_trips() {
    let mut workbook = Workbook::new();
    let custom = two_cell(8, 2, 12, 9);
    workbook.worksheet_mut(0).unwrap().add_drawing(
        DrawingObject::comment(1, 1, CellComment::new("a", "moved popup"))
            .with_anchor(custom.clone()),
    );

    let read = round_trip(&workbook);
    let comment = read
        .worksheet(0)
        .unwrap()
        .comments_drawn()
        .next()
        .expect("comment");
    assert_eq!((comment.row, comment.col), (1, 1));
    assert_eq!(comment.object.anchor, custom);
}

/// The comments part must use the MS-XLSB record ids Excel expects
/// (BrtBeginComments = 0x0274 per MS-XLSB 2.4.33; the worked example
/// in MS-XLSB 2.1.4 shows record type 637 = BrtCommentText). The
/// legacy emit was 0x0278-based, which Excel refuses to open.
// features: Plain-text comments with author
#[test]
fn xlsb_comments_part_uses_spec_record_ids() {
    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .set_comment_at(5, 5, CellComment::new("probe", "a comment"));
    let bytes = write_bytes(&workbook);

    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut part = Vec::new();
    {
        use std::io::Read;
        archive
            .by_name("xl/comments1.bin")
            .expect("comments part")
            .read_to_end(&mut part)
            .unwrap();
    }

    let mut ids = Vec::new();
    let mut i = 0usize;
    while i < part.len() {
        let mut id = u32::from(part[i]) & 0x7F;
        if part[i] & 0x80 != 0 {
            id |= (u32::from(part[i + 1]) & 0x7F) << 7;
            i += 2;
        } else {
            i += 1;
        }
        let mut size = 0u32;
        let mut shift = 0;
        loop {
            let byte = part[i];
            i += 1;
            size |= (u32::from(byte) & 0x7F) << shift;
            if byte & 0x80 == 0 {
                break;
            }
            shift += 7;
        }
        ids.push(id);
        i += size as usize;
    }

    assert_eq!(
        ids,
        vec![0x0274, 0x0276, 0x0278, 0x0277, 0x0279, 0x027B, 0x027D, 0x027C, 0x027A, 0x0275],
        "comments part record sequence must match MS-XLSB (BrtBeginComments .. BrtEndComments)"
    );
}

/// An anonymous (empty-author) comment keeps its own author slot
/// instead of being attributed to the first named author.
#[test]
fn xlsb_empty_author_comment_keeps_attribution() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_comment_at(0, 0, CellComment::new("Bob", "named"));
    sheet.set_comment_at(1, 0, CellComment::new("", "anonymous"));

    let read = round_trip(&workbook);
    let sheet = read.worksheet(0).unwrap();
    assert_eq!(sheet.comment_at(0, 0).unwrap().author, "Bob");
    assert_eq!(
        sheet.comment_at(1, 0).unwrap().author,
        "",
        "anonymous comment must not inherit another author"
    );
}
