//! Tests that PIN known XLS formula round-trip limitations. Each
//! test documents what survives, what doesn't, and references the
//! FEATURES.md notes column for the relevant row.
//!
//! The audit that landed this file found that intersection and
//! union formula text is silently dropped by the XLS writer →
//! reader round-trip, even though the cached value survives. The
//! FEATURES.md cells for those rows are R●/W● (partial) rather
//! than R✔/W✔. These tests assert the loss so a future fix to the
//! XLS formula compiler can be detected (the tests would flip from
//! "documents loss" to "fix landed, please update FEATURES.md").

use std::io::Cursor;

use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("write");
    XlsReader::read(Cursor::new(&bytes)).expect("read")
}

#[test]
fn intersection_formula_text_is_lost_through_xls_roundtrip() {
    // Pin the documented XLS limitation: the writer drops the
    // intersection operator from the emitted formula bytes, so the
    // reader returns no formula text for cells that used `=SUM(A B)`
    // style intersection formulas. The cached value (set explicitly
    // via set_formula_result) survives.
    //
    // If this assertion ever flips to `Some(...)`, the XLS formula
    // compiler has been fixed and FEATURES.md row 49 (Intersection)
    // should be updated from R●/W● to R✔/W✔ for the XLS column.
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
    assert_eq!(
        s.get_formula_at(0, 4),
        None,
        "documented XLS limitation: intersection formula text not \
         preserved; if this returns Some(...) the writer was fixed \
         and FEATURES.md should be updated"
    );
    // The cached value still survives, so applications that only
    // read the displayed cell value still work for these formulas.
    match s.get_value_at(0, 4).effective_value() {
        CellValue::Number(n) => assert!((n - 13.0).abs() < 1e-9),
        other => panic!("E1 expected Number(13), got {other:?}"),
    }
}

#[test]
fn union_formula_text_is_lost_through_xls_roundtrip() {
    // Same documented limitation as the intersection case above but
    // for the comma union operator inside bare parens. If this ever
    // returns Some(...), update FEATURES.md row 50 (Range union).
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
    assert_eq!(
        s.get_formula_at(0, 4),
        None,
        "documented XLS limitation: union formula text not preserved"
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
