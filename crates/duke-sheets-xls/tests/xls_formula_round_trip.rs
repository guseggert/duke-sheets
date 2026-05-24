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
use duke_sheets_xls::{biff, cfb::CompoundFile, XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("write");
    XlsReader::read(Cursor::new(&bytes)).expect("read")
}

fn only_formula_tokens(wb: &Workbook) -> Vec<u8> {
    let bytes = XlsWriter::write_to_bytes(wb).expect("write");
    let cfb = CompoundFile::open(Cursor::new(bytes)).expect("open CFB");
    let stream_path = if cfb.exists("/Workbook") {
        "/Workbook"
    } else {
        "/Book"
    };
    let stream = cfb.read_stream(stream_path).expect("read workbook stream");
    let records = biff::read_all_records(&mut Cursor::new(stream)).expect("read BIFF records");
    let mut token_streams = records
        .iter()
        .filter(|rec| rec.record_type == biff::records::FORMULA)
        .map(|rec| {
            assert!(rec.data.len() >= 22, "FORMULA record too short");
            let cce = u16::from_le_bytes([rec.data[20], rec.data[21]]) as usize;
            assert!(
                rec.data.len() >= 22 + cce,
                "FORMULA record token stream truncated"
            );
            rec.data[22..22 + cce].to_vec()
        })
        .collect::<Vec<_>>();

    assert_eq!(
        token_streams.len(),
        1,
        "expected exactly one FORMULA record"
    );
    token_streams.pop().unwrap()
}

#[test]
fn named_constant_formula_uses_reference_class_name_ptg() {
    let mut wb = Workbook::new();
    wb.define_name("MyTax", "0.07").unwrap();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=MyTax*2").unwrap();
    ws.set_formula_result(0, 0, CellValue::Number(0.14))
        .unwrap();

    let tokens = only_formula_tokens(&wb);
    assert_eq!(
        tokens.first().copied(),
        Some(0x23),
        "defined names stay R-class so range names survive SUM(); tokens={tokens:02X?}"
    );

    let parsed = write_then_read(&wb);
    let formula = parsed
        .worksheet(0)
        .unwrap()
        .get_formula_at(0, 0)
        .expect("named constant formula text must survive");
    assert!(
        formula.contains("MyTax"),
        "named constant lost from formula: {formula:?}"
    );
}

#[test]
fn named_range_intersection_uses_reference_class_name_ptgs() {
    let mut wb = Workbook::new();
    wb.define_name("LeftBlock", "A1:B3").unwrap();
    wb.define_name("RightBlock", "B2:C3").unwrap();
    let ws = wb.worksheet_mut(0).unwrap();
    for r in 0..3u32 {
        for c in 0..3u16 {
            ws.set_cell_value_at(r, c, (r * 3 + c as u32 + 1) as f64)
                .unwrap();
        }
    }
    ws.set_cell_formula("E1", "=SUM(LeftBlock RightBlock)")
        .unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(13.0))
        .unwrap();

    let tokens = only_formula_tokens(&wb);
    assert_eq!(
        tokens.first().copied(),
        Some(0x23),
        "left NameRef must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(5).copied(),
        Some(0x23),
        "right NameRef must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert!(
        tokens.contains(&0x0F),
        "intersection must emit PtgIsect (0x0F); tokens={tokens:02X?}"
    );

    let parsed = write_then_read(&wb);
    let formula = parsed
        .worksheet(0)
        .unwrap()
        .get_formula_at(0, 4)
        .expect("named-range intersection formula text must survive");
    assert!(
        formula.contains("LeftBlock") && formula.contains("RightBlock"),
        "named ranges lost from intersection formula: {formula:?}"
    );
}

#[test]
fn named_range_union_uses_reference_class_name_ptgs() {
    let mut wb = Workbook::new();
    wb.define_name("TopCells", "A1:A2").unwrap();
    wb.define_name("RightCells", "C2:C3").unwrap();
    let ws = wb.worksheet_mut(0).unwrap();
    for r in 0..3u32 {
        for c in 0..3u16 {
            ws.set_cell_value_at(r, c, (r * 3 + c as u32 + 1) as f64)
                .unwrap();
        }
    }
    ws.set_cell_formula("E1", "=SUM((TopCells,RightCells))")
        .unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(20.0))
        .unwrap();

    let tokens = only_formula_tokens(&wb);
    assert_eq!(
        tokens.first().copied(),
        Some(0x23),
        "left NameRef must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(5).copied(),
        Some(0x23),
        "right NameRef must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert!(
        tokens.contains(&0x10),
        "union must emit PtgUnion (0x10); tokens={tokens:02X?}"
    );

    let parsed = write_then_read(&wb);
    let formula = parsed
        .worksheet(0)
        .unwrap()
        .get_formula_at(0, 4)
        .expect("named-range union formula text must survive");
    assert!(
        formula.contains("TopCells") && formula.contains("RightCells"),
        "named ranges lost from union formula: {formula:?}"
    );
}

#[test]
fn range_operator_with_named_endpoints_emits_ptg_range() {
    let mut wb = Workbook::new();
    wb.define_name("StartCell", "A1").unwrap();
    wb.define_name("EndCell", "A3").unwrap();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).unwrap();
    ws.set_cell_value("A2", 2.0).unwrap();
    ws.set_cell_value("A3", 3.0).unwrap();
    ws.set_cell_formula("C1", "=SUM(StartCell:EndCell)")
        .unwrap();
    ws.set_formula_result(0, 2, CellValue::Number(6.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    assert_eq!(
        tokens.first().copied(),
        Some(0x23),
        "left endpoint must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(5).copied(),
        Some(0x23),
        "right endpoint must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert!(
        tokens.contains(&0x11),
        "range operator must emit PtgRange (0x11); tokens={tokens:02X?}"
    );

    let parsed = write_then_read(&wb);
    let formula = parsed
        .worksheet(0)
        .unwrap()
        .get_formula_at(0, 2)
        .expect("named-endpoint range formula text must survive");
    assert!(
        formula.contains("StartCell") && formula.contains("EndCell") && formula.contains(':'),
        "named endpoint range lost from formula: {formula:?}"
    );
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
