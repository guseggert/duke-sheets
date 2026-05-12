//! Round-trip tests for XLS formula expressions that exercise
//! reference operators: intersection (space), union (comma), range
//! (colon), cross-sheet refs, and named-range references.
//!
//! Originally landed during the cheat audit to PIN the loss of
//! intersection / union formula text through the XLS writer; the
//! follow-up fix (PtgIsect / PtgUnion / PtgRange emission with
//! R-class operands) flipped these to positive assertions.

use std::io::Cursor;

use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("write");
    XlsReader::read(Cursor::new(&bytes)).expect("read")
}

#[test]
fn intersection_formula_text_survives_xls_roundtrip() {
    // The XLS compiler emits PtgIsect (0x0F) for `=SUM(A1:B3 B2:C3)`
    // with R-class PtgArea operands so Excel can intersect the two
    // ranges to get cells B2 and B3 (5 + 8 = 13).
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for r in 0..3u32 {
        for c in 0..3u16 {
            ws.set_cell_value_at(r, c, (r * 3 + c as u32 + 1) as f64)
                .unwrap();
        }
    }
    ws.set_cell_formula("E1", "=SUM(A1:B3 B2:C3)").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(13.0))
        .unwrap();

    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    let f = s
        .get_formula_at(0, 4)
        .expect("intersection formula text must survive XLS round-trip");
    assert!(
        f.contains("A1:B3") && f.contains("B2:C3"),
        "intersection ranges lost from formula: {f:?}"
    );
    match s.get_value_at(0, 4).effective_value() {
        CellValue::Number(n) => assert!((n - 13.0).abs() < 1e-9),
        other => panic!("E1 expected Number(13), got {other:?}"),
    }
}

#[test]
fn union_formula_text_survives_xls_roundtrip() {
    // The XLS compiler emits PtgUnion (0x10) for `=SUM((A1:A2,C2:C3))`
    // with R-class PtgArea operands. SUM collects the four cells
    // {A1, A2, C2, C3} = 1+4+6+9 = 20.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for r in 0..3u32 {
        for c in 0..3u16 {
            ws.set_cell_value_at(r, c, (r * 3 + c as u32 + 1) as f64)
                .unwrap();
        }
    }
    ws.set_cell_formula("E1", "=SUM((A1:A2,C2:C3))").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(20.0))
        .unwrap();

    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    let f = s
        .get_formula_at(0, 4)
        .expect("union formula text must survive XLS round-trip");
    assert!(
        f.contains("A1:A2") && f.contains("C2:C3"),
        "union ranges lost from formula: {f:?}"
    );
    match s.get_value_at(0, 4).effective_value() {
        CellValue::Number(n) => assert!((n - 20.0).abs() < 1e-9),
        other => panic!("E1 expected Number(20), got {other:?}"),
    }
}

#[test]
fn cross_sheet_formula_text_survives_xls_roundtrip() {
    // Sanity check that the simpler features documented as R✔/W✔
    // really do round-trip. If this regresses we've broken something
    // unrelated to the intersection/union limitation.
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 5.0)
        .unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_formula("B1", "=Data!A1")
        .unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(5.0))
        .unwrap();

    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    let f = s
        .get_formula_at(0, 1)
        .expect("cross-sheet formula text must survive XLS round-trip");
    assert!(
        f.contains("Data") && f.contains("A1"),
        "cross-sheet reference lost: {f:?}"
    );
}

#[test]
fn named_range_in_formula_text_survives_xls_roundtrip() {
    // Per FEATURES.md row 204: "Names with formula bodies" is R●/W●
    // for XLS — the formula TEXT survives even though the workbook-
    // level `workbook.named_ranges()` map is not repopulated by the
    // reader. This test pins the formula-text half of that claim.
    let mut wb = Workbook::new();
    wb.define_name("MyTax", "0.07").unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 100.0)
        .unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_formula("B1", "=A1*MyTax")
        .unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(7.0))
        .unwrap();

    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    let f = s
        .get_formula_at(0, 1)
        .expect("named-range formula text must survive");
    assert!(
        f.contains("MyTax"),
        "named range MyTax lost from formula: {f:?}"
    );

    // Workbook-level named_ranges() is documented as NOT repopulated
    // by the XLS reader; the test pins this so we know if it ever
    // starts working.
    assert!(
        parsed.named_ranges().is_empty(),
        "XLS reader is documented as not repopulating workbook.named_ranges(); \
         got {:?} — if this is non-empty the reader has been improved and \
         FEATURES.md rows 202-205 should flip to R✔ for the XLS column",
        parsed
            .named_ranges()
            .iter()
            .map(|n| n.name.as_str())
            .collect::<Vec<_>>()
    );
}
