//! Round-trip tests for the XLS writer's FONT/XF emission
//! (slice 4a: font name, size, bold, italic, underline, strikethrough,
//! color via indexed palette / auto). Number format, fill, border, and
//! alignment are not yet emitted by the writer and round-trip to
//! defaults.

use std::io::Cursor;

use duke_sheets_core::style::{Color, FontVerticalAlign, Style, Underline};
use duke_sheets_core::{Workbook, Worksheet};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn font_at(sheet: &Worksheet, addr: &str) -> duke_sheets_core::style::FontStyle {
    sheet
        .cell_style(addr)
        .expect("cell_style ok")
        .map(|s| s.font.clone())
        .unwrap_or_default()
}

#[test]
fn bold_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let bold = Style::new().bold(true);
    ws.set_cell_style("A1", &bold).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(font_at(sheet, "A1").bold);
}

#[test]
fn italic_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "italic me").expect("set A1");
    let italic = Style::new().italic(true);
    ws.set_cell_style("A1", &italic).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(font_at(sheet, "A1").italic);
}

#[test]
fn font_size_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    let big = Style::new().font_size(18.0);
    ws.set_cell_style("A1", &big).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(font_at(sheet, "A1").size, 18.0);
}

#[test]
fn font_name_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "named").expect("set A1");
    let arial = Style::new().font_name("Arial");
    ws.set_cell_style("A1", &arial).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(font_at(sheet, "A1").name, "Arial");
}

#[test]
fn underline_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "underlined").expect("set A1");
    let mut style = Style::new();
    style.font.underline = Underline::Single;
    ws.set_cell_style("A1", &style).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(font_at(sheet, "A1").underline, Underline::Single);
}

#[test]
fn strikethrough_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "struck").expect("set A1");
    let mut style = Style::new();
    style.font.strikethrough = true;
    ws.set_cell_style("A1", &style).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(font_at(sheet, "A1").strikethrough);
}

#[test]
fn superscript_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "sup").expect("set A1");
    let mut style = Style::new();
    style.font.vertical_align = FontVerticalAlign::Superscript;
    ws.set_cell_style("A1", &style).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        font_at(sheet, "A1").vertical_align,
        FontVerticalAlign::Superscript
    );
}

#[test]
fn indexed_color_round_trips() {
    // Palette index 2 = red in the BIFF8 default 56-color palette.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "red").expect("set A1");
    let red = Style::new().font_color(Color::Indexed(2));
    ws.set_cell_style("A1", &red).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let color = font_at(sheet, "A1").color;
    let (r, g, b) = color.to_rgb();
    assert_eq!((r, g, b), (255, 0, 0), "got {color:?}");
}

#[test]
fn rgb_color_falls_back_to_auto() {
    // BIFF8 has no first-class RGB without a PALETTE record; the
    // writer falls back to Auto for arbitrary RGB. Document that
    // behavior so a future PALETTE-emission slice has a red test.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "purple").expect("set A1");
    let purple = Style::new().font_color(Color::Rgb {
        r: 128,
        g: 0,
        b: 128,
    });
    ws.set_cell_style("A1", &purple).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let color = font_at(sheet, "A1").color;
    assert!(
        matches!(color, Color::Auto | Color::Indexed(_)),
        "RGB falls back to Auto/Indexed in current writer; got {color:?}"
    );
}

#[test]
fn multiple_fonts_per_workbook_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    ws.set_cell_value("A2", 2.0).expect("set A2");
    ws.set_cell_value("A3", 3.0).expect("set A3");
    ws.set_cell_style("A1", &Style::new().bold(true))
        .expect("A1 bold");
    ws.set_cell_style("A2", &Style::new().italic(true))
        .expect("A2 italic");
    ws.set_cell_style("A3", &Style::new().font_size(20.0))
        .expect("A3 size");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(font_at(sheet, "A1").bold);
    assert!(font_at(sheet, "A2").italic);
    assert_eq!(font_at(sheet, "A3").size, 20.0);
}

#[test]
fn duplicate_styles_share_one_xf() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let bold = Style::new().bold(true);
    for row in 0..10u32 {
        let addr = format!("A{}", row + 1);
        ws.set_cell_value(&addr, row as f64).expect("set");
        ws.set_cell_style(&addr, &bold).expect("set style");
    }

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    for row in 0..10u32 {
        let addr = format!("A{}", row + 1);
        assert!(font_at(sheet, &addr).bold, "row {row}");
    }
}

#[test]
fn cells_without_style_remain_default() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("set A1");
    ws.set_cell_value("A2", "plain").expect("set A2");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let f1 = font_at(sheet, "A1");
    let f2 = font_at(sheet, "A2");
    assert!(!f1.bold && !f1.italic);
    assert!(!f2.bold && !f2.italic);
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_read_styled_cells_we_emit() {
    use std::net::TcpStream;
    use std::time::Duration;

    if TcpStream::connect_timeout(
        &"127.0.0.1:2002".parse().unwrap(),
        Duration::from_secs(2),
    )
    .is_err()
    {
        panic!("LO URP not reachable; start with `mise run urp:start`");
    }

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "bold").expect("set A1");
    ws.set_cell_value("B1", 99.0).expect("set B1");
    ws.set_cell_style("A1", &Style::new().bold(true)).expect("A1 style");
    ws.set_cell_style("B1", &Style::new().italic(true).font_size(14.0))
        .expect("B1 style");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_fonts_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<(String, f64), String> = rt.block_on(async {
        let mut bridge = duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect(
            "127.0.0.1",
            2002,
        )
        .await
        .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let a1 = wb.get_cell_string("A1").await.map_err(|e| format!("A1: {e}"))?;
        let b1 = wb.get_cell_value("B1").await.map_err(|e| format!("B1: {e}"))?;
        Ok((a1, b1))
    });
    let _ = std::fs::remove_file(&path);
    let (a1, b1) = outcome.expect("LO must read what we wrote");
    assert_eq!(a1, "bold");
    assert!((b1 - 99.0).abs() < 1e-9);
}
