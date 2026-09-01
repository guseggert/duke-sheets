//! Round-trip tests for the XLS writer's page setup records:
//! SETUP (0x00A1), HEADER (0x0014), FOOTER (0x0015),
//! LEFT/RIGHT/TOP/BOTTOM_MARGIN, PRINTHEADERS, PRINTGRIDLINES,
//! HPAGEBREAKS, VPAGEBREAKS.

use std::io::Cursor;

use duke_sheets_core::worksheet::{PageBreak, PageOrientation};
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn portrait_orientation_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.orientation = PageOrientation::Portrait;
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed.worksheet(0).unwrap().page_setup().orientation,
        PageOrientation::Portrait
    );
}

#[test]
fn landscape_orientation_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.orientation = PageOrientation::Landscape;
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed.worksheet(0).unwrap().page_setup().orientation,
        PageOrientation::Landscape
    );
}

#[test]
fn paper_size_a4_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.paper_size = 9; // A4
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.worksheet(0).unwrap().page_setup().paper_size, 9);
}

#[test]
fn scale_percentage_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.scale = 75;
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.worksheet(0).unwrap().page_setup().scale, 75);
}

#[test]
fn page_margins_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.left_margin = 1.25;
    ps.right_margin = 1.0;
    ps.top_margin = 1.5;
    ps.bottom_margin = 0.5;
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let ps = parsed.worksheet(0).unwrap().page_setup();
    assert!((ps.left_margin - 1.25).abs() < 1e-9);
    assert!((ps.right_margin - 1.0).abs() < 1e-9);
    assert!((ps.top_margin - 1.5).abs() < 1e-9);
    assert!((ps.bottom_margin - 0.5).abs() < 1e-9);
}

#[test]
fn header_footer_margins_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.header_margin = 0.4;
    ps.footer_margin = 0.6;
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let ps = parsed.worksheet(0).unwrap().page_setup();
    assert!((ps.header_margin - 0.4).abs() < 1e-9);
    assert!((ps.footer_margin - 0.6).abs() < 1e-9);
}

#[test]
fn odd_header_text_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.odd_header = Some("&LLeft&CCenter&RRight".to_string());
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed
            .worksheet(0)
            .unwrap()
            .page_setup()
            .odd_header
            .as_deref(),
        Some("&LLeft&CCenter&RRight")
    );
}

#[test]
fn header_formatting_codes_round_trip() {
    // BIFF8 stores header/footer text verbatim, so Excel-style
    // formatting codes (&B bold, &I italic, &"Arial,Bold" font face,
    // &14 size, &K00FF00 hex color) round-trip as raw substring.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.odd_header = Some("&L&BBold&I Italic&K00FF00 Green&\"Arial,Regular\"&14Bigger".to_string());
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let header = parsed
        .worksheet(0)
        .unwrap()
        .page_setup()
        .odd_header
        .clone()
        .expect("header present");
    assert!(header.contains("&B"), "got {header:?}");
    assert!(header.contains("&I"), "got {header:?}");
    assert!(header.contains("&K00FF00"), "got {header:?}");
    assert!(header.contains("&\"Arial,Regular\""), "got {header:?}");
    assert!(header.contains("&14"), "got {header:?}");
}

#[test]
fn odd_footer_text_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.odd_footer = Some("Page &P of &N".to_string());
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    assert_eq!(
        parsed
            .worksheet(0)
            .unwrap()
            .page_setup()
            .odd_footer
            .as_deref(),
        Some("Page &P of &N")
    );
}

#[test]
fn print_gridlines_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.print_gridlines = true;
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    assert!(parsed.worksheet(0).unwrap().page_setup().print_gridlines);
}

#[test]
fn print_headings_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.print_headings = true;
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    assert!(parsed.worksheet(0).unwrap().page_setup().print_headings);
}

#[test]
fn fit_to_pages_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.fit_to_width = Some(2);
    ps.fit_to_height = Some(3);
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let ps = parsed.worksheet(0).unwrap().page_setup();
    assert_eq!(ps.fit_to_width, Some(2));
    assert_eq!(ps.fit_to_height, Some(3));
}

#[test]
fn row_page_breaks_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_row_breaks(vec![
        PageBreak {
            id: 10,
            min: 0,
            max: 16383,
            man: true,
            pt: false,
        },
        PageBreak {
            id: 20,
            min: 0,
            max: 16383,
            man: true,
            pt: false,
        },
    ]);

    let parsed = write_then_read(&wb);
    let breaks = parsed.worksheet(0).unwrap().row_breaks();
    assert_eq!(breaks.len(), 2);
    assert_eq!(breaks[0].id, 10);
    assert_eq!(breaks[1].id, 20);
}

#[test]
fn col_page_breaks_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_col_breaks(vec![PageBreak {
        id: 5,
        min: 0,
        max: 65535,
        man: true,
        pt: false,
    }]);

    let parsed = write_then_read(&wb);
    let breaks = parsed.worksheet(0).unwrap().col_breaks();
    assert_eq!(breaks.len(), 1);
    assert_eq!(breaks[0].id, 5);
}

#[test]
fn full_page_setup_combination_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.orientation = PageOrientation::Landscape;
    ps.paper_size = 9;
    ps.scale = 90;
    ps.left_margin = 0.5;
    ps.right_margin = 0.5;
    ps.top_margin = 0.75;
    ps.bottom_margin = 0.75;
    ps.header_margin = 0.3;
    ps.footer_margin = 0.3;
    ps.odd_header = Some("Header".into());
    ps.odd_footer = Some("Footer".into());
    ps.print_gridlines = true;
    ps.print_headings = true;
    ws.set_page_setup(ps);

    let parsed = write_then_read(&wb);
    let ps = parsed.worksheet(0).unwrap().page_setup();
    assert_eq!(ps.orientation, PageOrientation::Landscape);
    assert_eq!(ps.paper_size, 9);
    assert_eq!(ps.scale, 90);
    assert!((ps.left_margin - 0.5).abs() < 1e-9);
    assert!((ps.top_margin - 0.75).abs() < 1e-9);
    assert!((ps.header_margin - 0.3).abs() < 1e-9);
    assert_eq!(ps.odd_header.as_deref(), Some("Header"));
    assert_eq!(ps.odd_footer.as_deref(), Some("Footer"));
    assert!(ps.print_gridlines);
    assert!(ps.print_headings);
}

/// LibreOffice must accept our SETUP / HEADER / FOOTER / margin /
/// PRINTHEADERS / PRINTGRIDLINES / HPAGEBREAKS / VPAGEBREAKS bundle.
/// Any one of these emitted with bad bytes (wrong endian, off
/// length, missing reserved fields) would cause LO's loader to flag
/// the file as corrupt or recompute a different page layout.
#[test]
fn lo_can_read_page_setup_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "page1").expect("A1");
    let mut ps = ws.page_setup().clone();
    ps.orientation = PageOrientation::Landscape;
    ps.left_margin = 0.5;
    ps.right_margin = 0.5;
    ps.top_margin = 0.75;
    ps.bottom_margin = 0.75;
    ps.header_margin = 0.3;
    ps.footer_margin = 0.3;
    ps.odd_header = Some("Page Header".into());
    ps.odd_footer = Some("Page &P of &N".into());
    ps.print_gridlines = true;
    ps.print_headings = true;
    ps.scale = 80;
    ws.set_page_setup(ps);
    ws.add_row_break(5);
    ws.add_col_break(3);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_pagesetup_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<String, String> = rt.block_on(async {
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
        wb.get_cell_string("A1")
            .await
            .map_err(|e| format!("A1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    let a1 = outcome.expect("LO must open page-setup workbook");
    assert_eq!(a1, "page1");
}
