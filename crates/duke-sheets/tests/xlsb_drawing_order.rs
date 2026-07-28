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
            part_rels: Vec::new(),
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
        .add_drawing(png("Pic").with_anchor(two_cell(1, 1, 3, 3))).unwrap();

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["image"]);
    let image = read.worksheet(0).unwrap().images().next().expect("image");
    assert_eq!(image.payload.data, PNG_1PX);
    assert_eq!(image.object.unwrap().meta.name.as_deref(), Some("Pic"));
    assert_eq!(image.object.unwrap().anchor, two_cell(1, 1, 3, 3));
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
    ).unwrap();
    sheet.add_drawing(
        png("Absolute").with_anchor(DrawingAnchor::Absolute {
            x_emu: 2_000_000,
            y_emu: 1_000_000,
            width_emu: 900_000,
            height_emu: 500_000,
        }),
    ).unwrap();

    let read = round_trip(&workbook);
    let images: Vec<_> = read.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 2);
    assert_eq!(
        images[0].object.unwrap().anchor,
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
        images[1].object.unwrap().anchor,
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
    sheet.add_chart(chart, two_cell(2, 2, 8, 12)).unwrap();

    let mut read = round_trip(&workbook);
    {
        let sheet = read.worksheet_mut(0).unwrap();
        let chart = sheet.charts_mut().next().expect("chart survived");
        chart.title = Some("Edited".to_string());
    }
    let again = round_trip(&read);
    let chart = again.worksheet(0).unwrap().charts().next().expect("chart");
    assert_eq!(chart.payload.title.as_deref(), Some("Edited"));
    assert_eq!(chart.object.unwrap().anchor, two_cell(2, 2, 8, 12));
}

/// Control z-position among native objects survives via the
/// com14:compatSp placeholder twin.
// features: Drawing z-order across kinds
#[test]
fn xlsb_control_z_order_between_images_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Middle")).with_anchor(two_cell(1, 1, 3, 3)),
    ).unwrap();
    sheet.add_drawing(png("Above").with_anchor(two_cell(2, 2, 4, 4))).unwrap();

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["image", "control", "image"]);
    let sheet = read.worksheet(0).unwrap();
    let control = sheet.form_controls().next().unwrap();
    assert_eq!(control.payload.caption_text().as_deref(), Some("Middle"));
    assert_eq!(control.object.unwrap().anchor, two_cell(1, 1, 3, 3));
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
    ).unwrap();

    let read = round_trip(&workbook);
    let control = read
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("control");
    assert_eq!(control.object.unwrap().meta.name.as_deref(), Some("Check Box 9"));
}

/// Comments keep their order relative to controls via the shared
/// legacy VML sequence.
#[test]
fn xlsb_comment_control_relative_order_round_trips() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("First")).with_anchor(two_cell(0, 0, 2, 2)),
    ).unwrap();
    sheet.add_drawing(DrawingObject::comment(
        4,
        4,
        CellComment::new("a", "between the controls"),
    )).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(checkbox("Second")).with_anchor(two_cell(6, 6, 8, 8)),
    ).unwrap();

    let read = round_trip(&workbook);
    assert_eq!(kind_tags(&read), vec!["control", "comment", "control"]);
    assert_eq!(
        read.worksheet(0).unwrap().comment_at(4, 4).unwrap().plain_text(),
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
    ).unwrap();

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
    ).unwrap();

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
        .set_comment_at(5, 5, CellComment::new("probe", "a comment")).unwrap();
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

/// Parse a BIFF12 part into (record id, payload) pairs.
fn biff12_records(part: &[u8]) -> Vec<(u32, Vec<u8>)> {
    let mut records = Vec::new();
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
        records.push((id, part[i..i + size as usize].to_vec()));
        i += size as usize;
    }
    records
}

/// Decode an XLWideString payload (u32 char count + UTF-16LE).
fn wide_str(payload: &[u8]) -> String {
    let n = u32::from_le_bytes(payload[0..4].try_into().unwrap()) as usize;
    let units: Vec<u16> = payload[4..4 + 2 * n]
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    String::from_utf16(&units).unwrap()
}

/// A sheet whose only form control lives inside a group must still
/// emit the legacy VML part, its relationship, and a BrtDrawing that
/// points at the drawing relationship (not the VML one). The reader
/// resolves parts by relationship type, so only a bytes-level check
/// of the rid pairing catches a misaligned BrtDrawing pointer.
// features: Grouped drawing objects
#[test]
fn xlsb_group_nested_control_keeps_drawing_and_vml_rels_aligned() {
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
                kind: DrawingKind::FormControl(FormControl::new(FormControlKind::Checkbox {
                    caption: "In group".into(),
                    state: duke_sheets::CheckState::Checked,
                    cell_link: None,
                    no_3d: false,
                })),
            }],
        })
        .with_anchor(two_cell(0, 0, 3, 3)),
    ).unwrap();
    let bytes = write_bytes(&workbook);

    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).unwrap();
    let read_part = |archive: &mut zip::ZipArchive<Cursor<Vec<u8>>>, name: &str| -> Vec<u8> {
        use std::io::Read;
        let mut part = Vec::new();
        archive
            .by_name(name)
            .unwrap_or_else(|_| panic!("part {name} missing"))
            .read_to_end(&mut part)
            .unwrap();
        part
    };

    let rels = String::from_utf8(read_part(
        &mut archive,
        "xl/worksheets/_rels/sheet1.bin.rels",
    ))
    .unwrap();
    let rel_type_of = |rid: &str| -> String {
        let marker = format!("Id=\"{rid}\"");
        let entry = rels
            .split("<Relationship ")
            .find(|chunk| chunk.contains(&marker))
            .unwrap_or_else(|| panic!("relationship {rid} missing in {rels}"));
        entry
            .split("Type=\"")
            .nth(1)
            .unwrap()
            .split('"')
            .next()
            .unwrap()
            .to_string()
    };

    let sheet_part = read_part(&mut archive, "xl/worksheets/sheet1.bin");
    let records = biff12_records(&sheet_part);
    let drawing_rid = records
        .iter()
        .find(|(id, _)| *id == 0x0226)
        .map(|(_, payload)| wide_str(payload))
        .expect("BrtDrawing record");
    let legacy_rid = records
        .iter()
        .find(|(id, _)| *id == 0x0227)
        .map(|(_, payload)| wide_str(payload))
        .expect("BrtLegacyDrawing record");

    assert!(
        rel_type_of(&drawing_rid).ends_with("/drawing"),
        "BrtDrawing must point at the drawing relationship"
    );
    assert!(
        rel_type_of(&legacy_rid).ends_with("/vmlDrawing"),
        "BrtLegacyDrawing must point at the vmlDrawing relationship"
    );
    // And the VML part itself carries the control shape.
    let vml = String::from_utf8(read_part(&mut archive, "xl/drawings/vmlDrawing1.vml")).unwrap();
    assert!(vml.contains("Checkbox"), "control shape emitted in VML");

    let read = XlsbReader::read(Cursor::new(bytes)).expect("read");
    let sheet = read.worksheet(0).unwrap();
    let controls = sheet.form_controls().collect::<Vec<_>>();
    assert_eq!(controls.len(), 1, "group-nested control survives");
}

/// A raw SmartArt anchor references its four parts through
/// `dgm:relIds` attributes (r:dm/r:lo/r:qs/r:cs) rather than
/// r:id/r:embed/r:link. Every attribute whose value matches a rel id
/// in the drawing's .rels must be captured, with target parts, and
/// survive a round trip resolvable.
#[test]
fn xlsb_smartart_rel_ids_attributes_are_captured_and_round_trip() {
    use std::io::{Read, Write};

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
    let bytes = write_bytes(&workbook);

    // Rebuild the package with the anchor spliced into drawing1.xml,
    // the diagram rels appended, and the dummy parts added.
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
            for (id, rel_type, target) in RT_DIAGRAM {
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
    for (path, content) in PARTS {
        let options = zip::write::SimpleFileOptions::default();
        out.start_file(path, options).unwrap();
        out.write_all(content.as_bytes()).unwrap();
    }
    let patched = out.finish().unwrap().into_inner();

    let assert_diagram_rels = |workbook: &Workbook, phase: &str| {
        let sheet = workbook.worksheet(0).unwrap();
        let raw = sheet
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

    let read = XlsbReader::read(Cursor::new(patched)).expect("read patched");
    assert_diagram_rels(&read, "first read");

    let again = round_trip(&read);
    assert_diagram_rels(&again, "round trip");
}

/// An anonymous (empty-author) comment keeps its own author slot
/// instead of being attributed to the first named author.
#[test]
fn xlsb_empty_author_comment_keeps_attribution() {
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

const WATERFALL_CHART_EX: &[u8] = br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$1:$A$3</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">a</cx:pt><cx:pt idx="1">b</cx:pt><cx:pt idx="2">c</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:f>Sheet1!$B$1:$B$3</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt><cx:pt idx="2">3</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Series1</cx:v></cx:txData></cx:tx><cx:dataId val="0"/><cx:layoutPr><cx:subtotals/></cx:layoutPr></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>"#;

/// Excel refuses a chartEx without chart style and chart colour style
/// sibling parts in XLSB exactly as in XLSX, so the XLSB writer must
/// also generate the defaults for a model-built chart.
#[test]
fn xlsb_model_built_chart_ex_gets_generated_style_and_color_parts() {
    use std::io::Read;

    let chart_ex =
        duke_sheets_chart::parse::parse_chart_ex_xml(WATERFALL_CHART_EX).expect("parse chartEx");
    assert!(chart_ex.raw_chart_style.is_none());
    assert!(chart_ex.raw_chart_color_style.is_none());

    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_drawing(DrawingObject::chart_ex(chart_ex).with_anchor(two_cell(0, 0, 4, 4)))
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut output).expect("write xlsb");
    let mut archive = zip::ZipArchive::new(Cursor::new(output.into_inner())).unwrap();

    let names: Vec<String> = archive.file_names().map(str::to_string).collect();
    for part in ["xl/charts/style1.xml", "xl/charts/colors1.xml"] {
        assert!(names.contains(&part.to_string()), "{part} missing: {names:?}");
    }

    let mut read = |name: &str| {
        let mut s = String::new();
        archive.by_name(name).unwrap().read_to_string(&mut s).unwrap();
        s
    };

    let content_types = read("[Content_Types].xml");
    for ct in [
        "application/vnd.ms-office.chartstyle+xml",
        "application/vnd.ms-office.chartcolorstyle+xml",
    ] {
        assert!(content_types.contains(ct), "content type {ct} missing");
    }

    let rels = read("xl/charts/_rels/chartEx1.xml.rels");
    for (kind, target) in [
        ("chartStyle", "style1.xml"),
        ("chartColorStyle", "colors1.xml"),
    ] {
        assert!(
            rels.contains(&format!("/relationships/{kind}\"")) && rels.contains(target),
            "chartEx rels missing {kind} -> {target}: {rels}"
        );
    }
}

/// Same contract as the XLSX writer: bytes Excel would reject fail the
/// write with an error naming the field and the defect.
#[test]
fn xlsb_write_rejects_a_raw_chart_style_excel_would_refuse() {
    let mut chart_ex =
        duke_sheets_chart::parse::parse_chart_ex_xml(WATERFALL_CHART_EX).expect("parse chartEx");
    chart_ex.raw_chart_style = Some(b"<cs:chartStyle/>".to_vec());

    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_drawing(DrawingObject::chart_ex(chart_ex).with_anchor(two_cell(0, 0, 4, 4)))
        .unwrap();

    let err = XlsbWriter::write(&workbook, Cursor::new(Vec::new()))
        .expect_err("garbage raw_chart_style must fail the write");
    let msg = err.to_string();
    assert!(
        msg.contains("raw_chart_style") && msg.contains("not bound"),
        "error must name the field and the defect: {msg}"
    );
}

/// Standard charts and chartEx charts are numbered independently, but
/// both families name their style parts styleN.xml/colorsN.xml, so the
/// chartEx numbers must continue above the standard charts' or the two
/// collide on style1.xml. Also pins that caller-supplied chartEx style
/// bytes are written verbatim while the missing colours part gets the
/// generated default.
#[test]
fn xlsb_standard_chart_and_chart_ex_style_parts_do_not_collide() {
    use std::io::Read;

    let standard_style = duke_sheets_chart::write::default_chart_style_bytes();
    let chart_ex_style = String::from_utf8(duke_sheets_chart::write::default_chart_style_bytes())
        .unwrap()
        .replace(r#" id="201""#, r#" id="999""#)
        .into_bytes();

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
    chart.raw_chart_style = Some(standard_style.clone());
    sheet.add_chart(chart, two_cell(2, 2, 8, 12)).unwrap();

    let mut chart_ex =
        duke_sheets_chart::parse::parse_chart_ex_xml(WATERFALL_CHART_EX).expect("parse chartEx");
    chart_ex.raw_chart_style = Some(chart_ex_style.clone());
    sheet
        .add_drawing(DrawingObject::chart_ex(chart_ex).with_anchor(two_cell(10, 2, 16, 12)))
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut output).expect("write xlsb");
    let mut archive = zip::ZipArchive::new(Cursor::new(output.into_inner())).unwrap();

    let mut read = |name: &str| -> Vec<u8> {
        let mut v = Vec::new();
        archive
            .by_name(name)
            .unwrap_or_else(|_| panic!("{name} missing"))
            .read_to_end(&mut v)
            .unwrap();
        v
    };

    assert_eq!(read("xl/charts/style1.xml"), standard_style, "standard chart style");
    assert_eq!(read("xl/charts/style2.xml"), chart_ex_style, "chartEx style must be verbatim");
    assert_eq!(
        read("xl/charts/colors2.xml"),
        duke_sheets_chart::write::default_chart_color_style_bytes(),
        "chartEx colours must be the generated default"
    );

    let rels = String::from_utf8(read("xl/charts/_rels/chartEx1.xml.rels")).unwrap();
    assert!(
        rels.contains(r#"Target="style2.xml""#) && rels.contains(r#"Target="colors2.xml""#),
        "chartEx rels must point at the offset numbers: {rels}"
    );
}

/// A drawing preserved verbatim keeps its chart part at the path its own
/// relationship names, so a chart the model writes must not be numbered
/// onto that path. Only image numbers used to be scanned for, so a
/// preserved chart1.xml and a model chart both claimed
/// xl/charts/chart1.xml and the write failed outright.
#[test]
fn xlsb_raw_preserved_chart_part_does_not_collide_with_a_model_chart() {
    use std::io::Read;

    use duke_sheets::{Chart, ChartType, DataReference, DataSeries};
    use duke_sheets_core::{RawDrawing, RawRel};

    let preserved_chart = b"<c:chartSpace/>".to_vec();
    let raw = RawDrawing {
        bytes: br#"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"#.to_vec(),
        rels: vec![RawRel {
            id: "rId1".into(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                .into(),
            target: "../charts/chart1.xml".into(),
            external: false,
            part: Some(preserved_chart.clone()),
            part_rels: Vec::new(),
        }],
    };

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(DrawingObject::raw(raw)).unwrap();
    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
    sheet.add_chart(chart, two_cell(6, 2, 12, 12)).unwrap();

    let mut output = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut output).expect("both charts must be writable");
    let mut archive = zip::ZipArchive::new(Cursor::new(output.into_inner())).unwrap();

    let mut read = |name: &str| -> Vec<u8> {
        let mut v = Vec::new();
        archive
            .by_name(name)
            .unwrap_or_else(|_| panic!("{name} missing"))
            .read_to_end(&mut v)
            .unwrap();
        v
    };
    assert_eq!(
        read("xl/charts/chart1.xml"),
        preserved_chart,
        "the preserved part must stay at the path its relationship names"
    );
    let model = String::from_utf8(read("xl/charts/chart2.xml")).unwrap();
    assert!(
        model.contains("Sheet1!$A$1:$A$3"),
        "the model chart must be numbered clear of it: {model}"
    );
}

/// A preserved part is not always self-contained. A chartEx reaches its
/// chart style and chart colour style parts through its own
/// relationships, and Excel refuses a workbook where those are missing,
/// so replaying the chartEx alone produced a file that would not open.
/// Diagrams have the same shape: their data part references the images.
#[test]
fn xlsb_preserved_part_keeps_its_own_relationships() {
    use std::io::Read;

    use duke_sheets_core::{RawDrawing, RawRel};

    let style = duke_sheets_chart::write::default_chart_style_bytes();
    let colors = duke_sheets_chart::write::default_chart_color_style_bytes();
    let chart_ex = b"<cx:chartSpace/>".to_vec();

    let raw = RawDrawing {
        bytes: br#"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><cx:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"#.to_vec(),
        rels: vec![RawRel {
            id: "rId1".into(),
            rel_type: "http://schemas.microsoft.com/office/2014/relationships/chartEx".into(),
            target: "../charts/chartEx1.xml".into(),
            external: false,
            part: Some(chart_ex.clone()),
            part_rels: vec![
                RawRel {
                    id: "rId1".into(),
                    rel_type: "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
                        .into(),
                    target: "style1.xml".into(),
                    external: false,
                    part: Some(style.clone()),
                    part_rels: Vec::new(),
                },
                RawRel {
                    id: "rId2".into(),
                    rel_type:
                        "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
                            .into(),
                    target: "colors1.xml".into(),
                    external: false,
                    part: Some(colors.clone()),
                    part_rels: Vec::new(),
                },
            ],
        }],
    };

    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_drawing(DrawingObject::raw(raw))
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut output).expect("write");
    let mut archive = zip::ZipArchive::new(Cursor::new(output.into_inner())).unwrap();

    let mut read = |name: &str| -> Vec<u8> {
        let mut v = Vec::new();
        archive
            .by_name(name)
            .unwrap_or_else(|_| panic!("{name} missing"))
            .read_to_end(&mut v)
            .unwrap();
        v
    };

    assert_eq!(read("xl/charts/chartEx1.xml"), chart_ex);
    assert_eq!(read("xl/charts/style1.xml"), style, "style part replayed");
    assert_eq!(read("xl/charts/colors1.xml"), colors, "colours part replayed");

    let rels = String::from_utf8(read("xl/charts/_rels/chartEx1.xml.rels")).unwrap();
    for target in ["style1.xml", "colors1.xml"] {
        assert!(
            rels.contains(target),
            "the preserved part's own relationships must be emitted: {rels}"
        );
    }

    let content_types = String::from_utf8(read("[Content_Types].xml")).unwrap();
    for ct in [
        "application/vnd.ms-office.chartstyle+xml",
        "application/vnd.ms-office.chartcolorstyle+xml",
    ] {
        assert!(content_types.contains(ct), "content type {ct} missing");
    }
}
