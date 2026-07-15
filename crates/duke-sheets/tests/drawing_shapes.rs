//! Basic worksheet shape coverage shared by XLSX, XLSB, and XLS.

use std::io::{Cursor, Read};

use duke_sheets::{
    CellMarker, CheckState, ChildTransform, Color, DrawingAnchor, DrawingKind, DrawingMeta,
    DrawingObject, DrawingText, EmbeddedImage, FormControl, FormControlKind, Group, GroupChild,
    GroupTransform, HorizontalAlignment, ImageFormat, RichTextRun, RunFont, Shape, ShapeFill,
    ShapeGeometry, ShapeLine, VerticalAlignment, Workbook, Worksheet,
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

fn basic_shape() -> Shape {
    let text = DrawingText {
        runs: vec![
            RichTextRun::with_font(
                "Bold ",
                RunFont {
                    name: Some("Segoe UI".to_string()),
                    size: Some(10.0),
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
            RichTextRun::with_font(
                "Italic",
                RunFont {
                    name: Some("Arial".to_string()),
                    size: Some(12.0),
                    italic: Some(true),
                    color: Some(Color::rgb(0, 0, 255)),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: Some(HorizontalAlignment::Center),
        vertical_alignment: Some(VerticalAlignment::Center),
    };

    Shape::rectangle()
        .with_fill(ShapeFill::Solid(Color::rgb(255, 0, 0)))
        .with_line(ShapeLine {
            color: Some(Color::rgb(0, 0, 255)),
            width_emu: Some(25_400),
            dash_style: Some("dash".to_string()),
            no_fill: false,
        })
        .with_text(text)
        .with_rotation(900_000)
        .with_flip_h(true)
}

fn basic_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    let index = sheet.add_shape(basic_shape(), anchor(1, 2, 5, 8)).unwrap();
    let object = &mut sheet.drawings_mut()[index];
    object.meta.name = Some("Status panel".to_string());
    object.meta.alt_text = Some("red status rectangle".to_string());
    object.meta.title = Some("Status".to_string());
    workbook
}

fn assert_basic_shape(sheet: &Worksheet, format: &str, has_title: bool) {
    assert_eq!(sheet.shape_count(), 1, "{format}: shape count");
    let drawn = sheet.shapes().next().expect("shape");
    assert_eq!(drawn.path, vec![0]);
    assert_eq!(drawn.object.unwrap().meta.name.as_deref(), Some("Status panel"));
    assert_eq!(
        drawn.object.unwrap().meta.alt_text.as_deref(),
        Some("red status rectangle")
    );
    assert_eq!(
        drawn.object.unwrap().meta.title.as_deref(),
        has_title.then_some("Status")
    );
    assert_eq!(drawn.object.unwrap().anchor, anchor(1, 2, 5, 8));

    let shape = drawn.payload;
    assert_eq!(shape.geometry, ShapeGeometry::Preset("rect".to_string()));
    assert_eq!(shape.fill, ShapeFill::Solid(Color::rgb(255, 0, 0)));
    assert_eq!(shape.line.color, Some(Color::rgb(0, 0, 255)));
    assert_eq!(shape.line.width_emu, Some(25_400));
    assert_eq!(shape.line.dash_style.as_deref(), Some("dash"));
    assert!(!shape.line.no_fill);
    assert_eq!(shape.rotation, 900_000);
    assert!(shape.flip_h);
    assert!(!shape.flip_v);

    let text = shape.text.as_ref().expect("shape text");
    assert_eq!(text.plain_text(), "Bold Italic");
    assert_eq!(text.horizontal_alignment, Some(HorizontalAlignment::Center));
    assert_eq!(text.vertical_alignment, Some(VerticalAlignment::Center));
    assert_eq!(text.runs.len(), 2);
    let bold = text.runs[0].font.as_ref().expect("bold font");
    assert_eq!(bold.name.as_deref(), Some("Segoe UI"));
    assert_eq!(bold.size, Some(10.0));
    assert_eq!(bold.bold, Some(true));
    let italic = text.runs[1].font.as_ref().expect("italic font");
    assert_eq!(italic.name.as_deref(), Some("Arial"));
    assert_eq!(italic.size, Some(12.0));
    assert_eq!(italic.italic, Some(true));
    assert_eq!(italic.color, Some(Color::rgb(0, 0, 255)));
}

fn round_trip_xlsx(workbook: &Workbook) -> (Workbook, Vec<u8>) {
    let mut output = Cursor::new(Vec::new());
    XlsxWriter::write(workbook, &mut output).expect("write xlsx");
    let bytes = output.into_inner();
    let parsed = XlsxReader::read(Cursor::new(&bytes)).expect("read xlsx");
    (parsed, bytes)
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

// features: Shapes (rectangles, arrows, ...); Text boxes; Drawing name / alternative text / title
#[test]
fn xlsx_basic_shape_round_trips() {
    let (parsed, _) = round_trip_xlsx(&basic_workbook());
    assert_basic_shape(parsed.worksheet(0).unwrap(), "xlsx", true);
}

// features: Shapes (rectangles, arrows, ...); Text boxes; Drawing name / alternative text / title
#[test]
fn xlsb_basic_shape_round_trips() {
    let parsed = round_trip_xlsb(&basic_workbook());
    assert_basic_shape(parsed.worksheet(0).unwrap(), "xlsb", true);
}

// features: Shapes (rectangles, arrows, ...); Text boxes; Drawing name / alternative text / title
#[test]
fn xls_basic_shape_round_trips() {
    let parsed = round_trip_xls(&basic_workbook());
    assert_basic_shape(parsed.worksheet(0).unwrap(), "xls", false);
}

// features: Grouped drawing objects
#[test]
fn xlsx_grouped_shape_child_round_trips() {
    let child = GroupChild {
        meta: DrawingMeta {
            name: Some("Child rectangle".to_string()),
            alt_text: Some("inside a group".to_string()),
            ..DrawingMeta::default()
        },
        transform: ChildTransform {
            x_emu: 100_000,
            y_emu: 200_000,
            cx_emu: 500_000,
            cy_emu: 300_000,
            ..ChildTransform::default()
        },
        kind: DrawingKind::Shape(Box::new(basic_shape())),
    };
    let group = Group {
        transform: GroupTransform {
            cx_emu: 1_000_000,
            cy_emu: 1_000_000,
            child_cx_emu: 1_000_000,
            child_cy_emu: 1_000_000,
            ..GroupTransform::default()
        },
        children: vec![child],
    };
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(
        DrawingObject::group(group)
            .with_anchor(anchor(0, 0, 4, 8))
            .with_name("Shape group"),
    ).unwrap();

    let (parsed, _) = round_trip_xlsx(&workbook);
    let group = parsed.worksheet(0).unwrap().drawings()[0]
        .kind
        .as_group()
        .expect("group");
    assert_eq!(group.children.len(), 1);
    assert_eq!(
        group.children[0].meta.name.as_deref(),
        Some("Child rectangle")
    );
    let shape = group.children[0].kind.as_shape().expect("shape child");
    assert_eq!(shape.geometry, ShapeGeometry::Preset("rect".to_string()));
    assert_eq!(shape.fill, ShapeFill::Solid(Color::rgb(255, 0, 0)));
    assert_eq!(shape.text.as_ref().unwrap().plain_text(), "Bold Italic");
}

fn image() -> EmbeddedImage {
    EmbeddedImage {
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
    }
}

fn z_order_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_image(image(), anchor(0, 0, 2, 2)).unwrap();
    sheet.add_shape(Shape::preset("ellipse"), anchor(1, 1, 3, 3)).unwrap();
    sheet.add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "top".into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: false,
        }),
        anchor(2, 2, 4, 4),
    ).unwrap();
    workbook
}

fn assert_shape_z_order(sheet: &Worksheet, format: &str) {
    let tags: Vec<_> = sheet
        .drawings()
        .iter()
        .map(|object| match object.kind {
            DrawingKind::Image(_) => "image",
            DrawingKind::Shape(_) => "shape",
            DrawingKind::FormControl(_) => "control",
            _ => "other",
        })
        .collect();
    assert_eq!(tags, ["image", "shape", "control"], "{format}: z-order");
}

#[test]
fn shape_z_order_round_trips_in_all_formats() {
    let workbook = z_order_workbook();
    assert_shape_z_order(round_trip_xlsx(&workbook).0.worksheet(0).unwrap(), "xlsx");
    assert_shape_z_order(round_trip_xlsb(&workbook).worksheet(0).unwrap(), "xlsb");
    assert_shape_z_order(round_trip_xls(&workbook).worksheet(0).unwrap(), "xls");
}

#[test]
fn xlsx_shape_preserves_unmodeled_shape_property_fragments() {
    let mut shape = Shape::rectangle();
    shape.raw_shape_properties = Some(
        br#"<a:effectLst><a:outerShdw blurRad="40000"><a:srgbClr val="00FF00"/></a:outerShdw></a:effectLst><a:extLst><a:ext uri="{D4A5E4CE-4971-4E86-8E55-9D104D6F43E1}"><probe:marker xmlns:probe="urn:duke-shape-test"/></a:ext></a:extLst>"#.to_vec(),
    );
    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_shape(shape, anchor(0, 0, 3, 5)).unwrap();

    let (first_read, _) = round_trip_xlsx(&workbook);
    let parsed_shape = first_read
        .worksheet(0)
        .unwrap()
        .shapes()
        .next()
        .unwrap()
        .payload;
    let raw = std::str::from_utf8(parsed_shape.raw_shape_properties.as_deref().unwrap()).unwrap();
    assert!(raw.contains("outerShdw"));
    assert!(raw.contains("urn:duke-shape-test"));

    let (_, second_bytes) = round_trip_xlsx(&first_read);
    let mut archive = zip::ZipArchive::new(Cursor::new(second_bytes)).unwrap();
    let mut drawing_xml = String::new();
    archive
        .by_name("xl/drawings/drawing1.xml")
        .unwrap()
        .read_to_string(&mut drawing_xml)
        .unwrap();
    assert!(drawing_xml.contains("outerShdw"));
    assert!(drawing_xml.contains("urn:duke-shape-test"));
}

/// Malformed preserved raw XML must fail the write instead of being
/// spliced verbatim into the drawing part.
#[test]
fn malformed_preserved_shape_xml_fails_the_write() {
    let mut shape = Shape::rectangle();
    shape.raw_shape_properties = Some(b"<a:effectLst><unclosed>".to_vec());
    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_shape(shape, anchor(0, 0, 2, 2)).unwrap();
    let error = XlsxWriter::write(&workbook, &mut Cursor::new(Vec::new())).unwrap_err();
    assert!(
        error.to_string().contains("unterminated"),
        "expected raw-XML validation error, got: {error}"
    );

    let mut shape = Shape::rectangle().with_text(DrawingText::from("t"));
    shape.raw_text_body = Some(b"<xdr:txBody><a:p></xdr:txBody><script/>".to_vec());
    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_shape(shape, anchor(0, 0, 2, 2)).unwrap();
    assert!(XlsxWriter::write(&workbook, &mut Cursor::new(Vec::new())).is_err());
}

/// Shape equality tracks emission semantics: a stale preserved
/// geometry (regenerates) is not equal to a current one (replays).
#[test]
fn shape_equality_tracks_preservation_staleness() {
    let raw = br#"<a:custGeom><a:pathLst/></a:custGeom>"#.to_vec();
    let mut preserved = Shape::preset("rect");
    preserved.set_preserved_shape_properties(Some(raw.clone()));

    let mut stale = preserved.clone();
    stale.set_geometry(ShapeGeometry::Preset("ellipse".to_string()));
    stale.set_geometry(ShapeGeometry::Preset("rect".to_string()));
    assert_eq!(
        preserved, stale,
        "matching modeled state and staleness outcome compare equal"
    );

    let mut edited = preserved.clone();
    edited.set_geometry(ShapeGeometry::Preset("ellipse".to_string()));
    assert_ne!(
        preserved, edited,
        "different modeled geometry compares unequal"
    );

    let mut caller_supplied = Shape::preset("rect");
    caller_supplied.raw_shape_properties = Some(raw);
    let mut snapshot_stale = caller_supplied.clone();
    snapshot_stale.raw_geometry_snapshot = Some(ShapeGeometry::Preset("ellipse".to_string()));
    assert_ne!(
        caller_supplied, snapshot_stale,
        "same bytes but different staleness emit different XML and must not compare equal"
    );
}
