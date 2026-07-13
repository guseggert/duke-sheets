//! Generate populated test fixtures for the language-binding test
//! suites: `cargo run -p duke-sheets --example gen_binding_fixtures -- <dir>`.
//!
//! The bindings expose comments, autofilters, data validations, and
//! embedded images as read-only APIs, so their tests cannot author
//! these features themselves; this writes `sample.xlsx`, `sample.xls`,
//! and `sample.xlsb` containing one of each at test time (binary
//! fixtures are never committed).

use duke_sheets::{Workbook, WorkbookExt};
use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, ValueFilter};
use duke_sheets_core::comment::CellComment;
use duke_sheets_core::named_range::{NameScope, NamedRange};
use duke_sheets_core::validation::DataValidation;
use duke_sheets_core::{CellAddress, CellRange, CellValue};

/// A 68-byte 1x1 transparent PNG with valid chunk CRCs.
const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82,
];

fn build(with_image: bool) -> Workbook {
    let mut wb = Workbook::new();
    wb.named_ranges_mut().define_or_update(NamedRange::new(
        "MyRange",
        "Sheet1!$A$2:$A$4",
        NameScope::Workbook,
    ));

    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Score").unwrap();
    ws.set_cell_value("A2", 1.0).unwrap();
    ws.set_cell_value("A3", 2.0).unwrap();
    ws.set_cell_value("A4", 3.0).unwrap();
    ws.set_cell_formula("B1", "=SUM(MyRange)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(6.0)).unwrap();

    ws.set_comment(
        "A1",
        CellComment {
            author: "Tester".to_string(),
            text: "fixture comment".to_string(),
        },
    )
    .unwrap();

    let mut af = AutoFilter::new(CellRange::new(
        CellAddress::parse("A1").unwrap(),
        CellAddress::parse("A4").unwrap(),
    ));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Values(ValueFilter {
            values: vec!["1".to_string(), "3".to_string()],
            blank: false,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let mut dv = DataValidation::list("Red,Green,Blue");
    dv.ranges = vec![CellRange::parse("C1:C5").unwrap()];
    ws.add_data_validation(dv);

    if with_image {
        let image = duke_sheets_chart::EmbeddedImage {
            format: duke_sheets_chart::ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 0,
            height_emu: 0,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        };
        let anchor = duke_sheets_chart::DrawingAnchor::TwoCell {
            from: duke_sheets_chart::CellMarker {
                col: 4,
                col_offset_emu: 0,
                row: 1,
                row_offset_emu: 0,
            },
            to: duke_sheets_chart::CellMarker {
                col: 6,
                col_offset_emu: 0,
                row: 4,
                row_offset_emu: 0,
            },
            edit_as: None,
        };
        ws.add_drawing(
            duke_sheets_core::DrawingObject::image(image)
                .with_anchor(anchor)
                .with_name("FixturePic"),
        );
    }

    wb
}

fn main() {
    let dir = std::env::args()
        .nth(1)
        .expect("usage: gen_binding_fixtures <output-dir>");
    std::fs::create_dir_all(&dir).expect("create output dir");

    // XLSB has no image read support yet; the others carry one.
    for (ext, with_image) in [("xlsx", true), ("xls", true), ("xlsb", false)] {
        let path = format!("{dir}/sample.{ext}");
        build(with_image).save(&path).expect("save fixture");
        println!("{path}");
    }
}
