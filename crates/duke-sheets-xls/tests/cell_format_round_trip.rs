//! Round-trip tests for the XLS writer's alignment / border / fill
//! XF axes (slice 4c).

use std::io::Cursor;

use duke_sheets_core::style::{
    BorderEdge, BorderLineStyle, Color, DiagonalDirection, FillStyle, HorizontalAlignment,
    PatternType, ReadingOrder, Style, VerticalAlignment,
};
use duke_sheets_core::{Workbook, Worksheet};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn style_at(sheet: &Worksheet, addr: &str) -> Style {
    sheet
        .cell_style(addr)
        .expect("cell_style ok")
        .cloned()
        .unwrap_or_default()
}

#[test]
fn horizontal_alignment_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let cases = [
        ("A1", HorizontalAlignment::Left),
        ("A2", HorizontalAlignment::Center),
        ("A3", HorizontalAlignment::Right),
        ("A4", HorizontalAlignment::Justify),
    ];
    for (addr, h) in &cases {
        ws.set_cell_value(addr, "x").expect("set");
        let style = Style::new().horizontal_alignment(*h);
        ws.set_cell_style(addr, &style).expect("set style");
    }

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    for (addr, h) in &cases {
        assert_eq!(style_at(sheet, addr).alignment.horizontal, *h, "{addr}");
    }
}

#[test]
fn vertical_alignment_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let cases = [
        ("A1", VerticalAlignment::Top),
        ("A2", VerticalAlignment::Center),
        ("A3", VerticalAlignment::Bottom),
    ];
    for (addr, v) in &cases {
        ws.set_cell_value(addr, "x").expect("set");
        let style = Style::new().vertical_alignment(*v);
        ws.set_cell_style(addr, &style).expect("set style");
    }

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    for (addr, v) in &cases {
        assert_eq!(style_at(sheet, addr).alignment.vertical, *v, "{addr}");
    }
}

#[test]
fn wrap_text_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "wrapped text").expect("set A1");
    let style = Style::new().wrap_text(true);
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(style_at(sheet, "A1").alignment.wrap_text);
}

#[test]
fn indent_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "indented").expect("set A1");
    let mut style = Style::new();
    style.alignment.indent = 3;
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(style_at(sheet, "A1").alignment.indent, 3);
}

#[test]
fn shrink_to_fit_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "shrink me to fit").expect("set A1");
    let mut style = Style::new();
    style.alignment.shrink_to_fit = true;
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(style_at(sheet, "A1").alignment.shrink_to_fit);
}

#[test]
fn rotation_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "rotated").expect("set A1");
    let mut style = Style::new();
    style.alignment.rotation = 45;
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(style_at(sheet, "A1").alignment.rotation, 45);
}

#[test]
fn reading_order_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "rtl").expect("set A1");
    ws.set_cell_value("A2", "ltr").expect("set A2");
    let mut rtl = Style::new();
    rtl.alignment.reading_order = ReadingOrder::RightToLeft;
    let mut ltr = Style::new();
    ltr.alignment.reading_order = ReadingOrder::LeftToRight;
    ws.set_cell_style("A1", &rtl).expect("A1 style");
    ws.set_cell_style("A2", &ltr).expect("A2 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        style_at(sheet, "A1").alignment.reading_order,
        ReadingOrder::RightToLeft
    );
    assert_eq!(
        style_at(sheet, "A2").alignment.reading_order,
        ReadingOrder::LeftToRight
    );
}

#[test]
fn border_thin_all_sides_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let mut style = Style::new();
    style.border =
        duke_sheets_core::style::BorderStyle::all(BorderLineStyle::Thin, Color::Indexed(0));
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let border = &style_at(sheet, "A1").border;
    let expect_edge = BorderEdge::new(BorderLineStyle::Thin, Color::Rgb { r: 0, g: 0, b: 0 });
    assert_eq!(border.left.as_ref(), Some(&expect_edge));
    assert_eq!(border.right.as_ref(), Some(&expect_edge));
    assert_eq!(border.top.as_ref(), Some(&expect_edge));
    assert_eq!(border.bottom.as_ref(), Some(&expect_edge));
}

#[test]
fn border_individual_sides_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let mut style = Style::new();
    style.border.left = Some(BorderEdge::new(BorderLineStyle::Thin, Color::Auto));
    style.border.right = Some(BorderEdge::new(BorderLineStyle::Medium, Color::Auto));
    style.border.top = Some(BorderEdge::new(BorderLineStyle::Thick, Color::Auto));
    style.border.bottom = Some(BorderEdge::new(BorderLineStyle::Double, Color::Auto));
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let border = &style_at(sheet, "A1").border;
    assert_eq!(
        border.left.as_ref().map(|e| e.style),
        Some(BorderLineStyle::Thin)
    );
    assert_eq!(
        border.right.as_ref().map(|e| e.style),
        Some(BorderLineStyle::Medium)
    );
    assert_eq!(
        border.top.as_ref().map(|e| e.style),
        Some(BorderLineStyle::Thick)
    );
    assert_eq!(
        border.bottom.as_ref().map(|e| e.style),
        Some(BorderLineStyle::Double)
    );
}

#[test]
fn border_color_indexed_non_black_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let mut style = Style::new();
    style.border.left = Some(BorderEdge::new(BorderLineStyle::Thin, Color::Indexed(2)));
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let edge = style_at(sheet, "A1")
        .border
        .left
        .clone()
        .expect("left edge present");
    let (r, g, b) = edge.color.to_rgb();
    assert_eq!(
        (r, g, b),
        (255, 0, 0),
        "indexed-2 (red) border color should round-trip; got {:?}",
        edge.color
    );
}

#[test]
fn diagonal_border_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let mut style = Style::new();
    style.border.diagonal = Some(BorderEdge::new(BorderLineStyle::Thin, Color::Auto));
    style.border.diagonal_direction = DiagonalDirection::Both;
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let border = &style_at(sheet, "A1").border;
    assert_eq!(
        border.diagonal.as_ref().map(|e| e.style),
        Some(BorderLineStyle::Thin)
    );
    assert_eq!(border.diagonal_direction, DiagonalDirection::Both);
}

#[test]
fn solid_fill_with_indexed_color_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let mut style = Style::new();
    // Index 5 in the BIFF8 default palette (icv 0x0D) = yellow.
    style.fill = FillStyle::Solid {
        color: Color::Indexed(5),
    };
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let fill = &style_at(sheet, "A1").fill;
    let is_solid_equivalent = match fill {
        FillStyle::Solid { .. } => true,
        FillStyle::Pattern { pattern, .. } => matches!(pattern, PatternType::Solid),
        _ => false,
    };
    assert!(
        is_solid_equivalent,
        "expected Solid-equivalent fill on round-trip, got {fill:?}"
    );
}

// features: Fill: solid color
#[test]
fn fill_color_palette_exact_rgb_round_trips() {
    // Palette-exact RGB must map to its default-palette icv the way
    // the font path does (violet = palette slot 12 = icv 20), not
    // drop to the 0x40 system default that reads back as black.
    let violet = Color::Rgb {
        r: 128,
        g: 0,
        b: 128,
    };
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let mut style = Style::new();
    style.fill = FillStyle::Solid { color: violet };
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let fill = &style_at(sheet, "A1").fill;
    let color = match fill {
        FillStyle::Solid { color } => *color,
        FillStyle::Pattern { foreground, .. } => *foreground,
        other => panic!("expected a solid-equivalent fill, got {other:?}"),
    };
    assert_eq!(
        color.to_rgb(),
        (128, 0, 128),
        "palette-exact fill RGB survives; got {color:?}"
    );
}

// features: Border: color
#[test]
fn border_color_palette_exact_rgb_round_trips() {
    // Light Orange (255,153,0) = default palette slot 44 = icv 52.
    let light_orange = Color::Rgb {
        r: 255,
        g: 153,
        b: 0,
    };
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let mut style = Style::new();
    style.border.left = Some(BorderEdge::new(BorderLineStyle::Thin, light_orange));
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let edge = style_at(sheet, "A1")
        .border
        .left
        .clone()
        .expect("left edge present");
    assert_eq!(
        edge.color.to_rgb(),
        (255, 153, 0),
        "palette-exact border RGB survives; got {:?}",
        edge.color
    );
}

#[test]
fn pattern_fill_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let mut style = Style::new();
    style.fill = FillStyle::Pattern {
        pattern: PatternType::DarkGrid,
        foreground: Color::Indexed(0),
        background: Color::Indexed(1),
    };
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let fill = &style_at(sheet, "A1").fill;
    match fill {
        FillStyle::Pattern { pattern, .. } => {
            assert_eq!(*pattern, PatternType::DarkGrid);
        }
        other => panic!("expected Pattern fill, got {other:?}"),
    }
}

#[test]
fn alignment_combined_with_font_and_fill_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "ALL THE STYLES").expect("set A1");

    let mut style = Style::new()
        .bold(true)
        .horizontal_alignment(HorizontalAlignment::Center)
        .vertical_alignment(VerticalAlignment::Center)
        .wrap_text(true);
    style.fill = FillStyle::Solid {
        color: Color::Indexed(2),
    };
    style.border = duke_sheets_core::style::BorderStyle::all(BorderLineStyle::Medium, Color::Auto);
    ws.set_cell_style("A1", &style).expect("set style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let s = style_at(sheet, "A1");
    assert!(s.font.bold);
    assert_eq!(s.alignment.horizontal, HorizontalAlignment::Center);
    assert_eq!(s.alignment.vertical, VerticalAlignment::Center);
    assert!(s.alignment.wrap_text);
    assert!(matches!(
        s.fill,
        FillStyle::Solid { .. } | FillStyle::Pattern { .. }
    ));
    let border = &s.border;
    assert!(matches!(
        border.left.as_ref().map(|e| e.style),
        Some(BorderLineStyle::Medium)
    ));
}

#[test]
fn distinct_alignment_styles_dedupe_xfs() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let center = Style::new().horizontal_alignment(HorizontalAlignment::Center);
    for row in 0..50u32 {
        let addr = format!("A{}", row + 1);
        ws.set_cell_value(&addr, row as f64).expect("set");
        ws.set_cell_style(&addr, &center).expect("set style");
    }

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    for row in 0..50u32 {
        let addr = format!("A{}", row + 1);
        assert_eq!(
            style_at(sheet, &addr).alignment.horizontal,
            HorizontalAlignment::Center,
            "row {row}"
        );
    }
}
