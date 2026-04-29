//! Round-trip tests for the XLS writer's FORMAT/XF.ifmt emission
//! (slice 4b: built-in number format indices and user-defined custom
//! format strings).

use std::io::Cursor;

use duke_sheets_core::style::{NumberFormat, Style};
use duke_sheets_core::{Workbook, Worksheet};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn number_format_at(sheet: &Worksheet, addr: &str) -> NumberFormat {
    sheet
        .cell_style(addr)
        .expect("cell_style ok")
        .map(|s| s.number_format.clone())
        .unwrap_or_default()
}

#[test]
fn builtin_percent_format_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 0.5).expect("set A1");
    let mut style = Style::new();
    style.number_format = NumberFormat::BuiltIn(NumberFormat::ID_PERCENT_DEC2);
    ws.set_cell_style("A1", &style).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        number_format_at(sheet, "A1"),
        NumberFormat::BuiltIn(NumberFormat::ID_PERCENT_DEC2)
    );
}

#[test]
fn builtin_date_format_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 45000.0).expect("set A1");
    let mut style = Style::new();
    style.number_format = NumberFormat::BuiltIn(NumberFormat::ID_DATE_SHORT);
    ws.set_cell_style("A1", &style).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        number_format_at(sheet, "A1"),
        NumberFormat::BuiltIn(NumberFormat::ID_DATE_SHORT)
    );
}

#[test]
fn builtin_number_format_with_thousands_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1234567.89).expect("set A1");
    let mut style = Style::new();
    style.number_format = NumberFormat::BuiltIn(NumberFormat::ID_NUMBER_SEP_DEC2);
    ws.set_cell_style("A1", &style).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        number_format_at(sheet, "A1"),
        NumberFormat::BuiltIn(NumberFormat::ID_NUMBER_SEP_DEC2)
    );
}

#[test]
fn custom_currency_format_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 99.95).expect("set A1");
    let style = Style::new().number_format("$#,##0.00");
    ws.set_cell_style("A1", &style).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        number_format_at(sheet, "A1"),
        NumberFormat::Custom("$#,##0.00".to_string())
    );
}

#[test]
fn custom_iso_date_format_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 45000.0).expect("set A1");
    let style = Style::new().number_format("yyyy-mm-dd");
    ws.set_cell_style("A1", &style).expect("set A1 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        number_format_at(sheet, "A1"),
        NumberFormat::Custom("yyyy-mm-dd".to_string())
    );
}

#[test]
fn duplicate_custom_formats_share_one_format_record() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let style = Style::new().number_format("0.0%");
    for row in 0..20u32 {
        let addr = format!("A{}", row + 1);
        ws.set_cell_value(&addr, row as f64 * 0.05)
            .expect("set value");
        ws.set_cell_style(&addr, &style).expect("set style");
    }

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    for row in 0..20u32 {
        let addr = format!("A{}", row + 1);
        assert_eq!(
            number_format_at(sheet, &addr),
            NumberFormat::Custom("0.0%".to_string()),
            "row {row}"
        );
    }
}

#[test]
fn multiple_distinct_formats_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 0.5).expect("A1");
    ws.set_cell_value("A2", 1234.5).expect("A2");
    ws.set_cell_value("A3", 99.95).expect("A3");

    let pct = Style::new().number_format("0.00%");
    let acct = Style::new().number_format("#,##0.00");
    let custom_currency = Style::new().number_format("[$$-409]#,##0.00");

    ws.set_cell_style("A1", &pct).expect("A1 style");
    ws.set_cell_style("A2", &acct).expect("A2 style");
    ws.set_cell_style("A3", &custom_currency).expect("A3 style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert_eq!(
        number_format_at(sheet, "A1"),
        NumberFormat::Custom("0.00%".to_string())
    );
    assert_eq!(
        number_format_at(sheet, "A2"),
        NumberFormat::Custom("#,##0.00".to_string())
    );
    assert_eq!(
        number_format_at(sheet, "A3"),
        NumberFormat::Custom("[$$-409]#,##0.00".to_string())
    );
}

#[test]
fn format_combined_with_font_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 0.25).expect("A1");
    let style = Style::new().bold(true).number_format("0.0%");
    ws.set_cell_style("A1", &style).expect("style");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let cell_style = sheet.cell_style("A1").unwrap().expect("style present");
    assert!(cell_style.font.bold);
    assert_eq!(
        cell_style.number_format,
        NumberFormat::Custom("0.0%".to_string())
    );
}

#[test]
fn cells_without_format_round_trip_to_general() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(matches!(
        number_format_at(sheet, "A1"),
        NumberFormat::General
    ));
}
