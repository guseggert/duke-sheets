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

/// Token bytes for the FORMULA record at a specific (row, col). Lets a test
/// place multiple formulas in one workbook and inspect each independently.
fn only_formula_tokens_at(wb: &Workbook, row: u16, col: u16) -> Vec<u8> {
    let bytes = XlsWriter::write_to_bytes(wb).expect("write");
    let cfb = CompoundFile::open(Cursor::new(bytes)).expect("open CFB");
    let stream_path = if cfb.exists("/Workbook") {
        "/Workbook"
    } else {
        "/Book"
    };
    let stream = cfb.read_stream(stream_path).expect("read workbook stream");
    let records = biff::read_all_records(&mut Cursor::new(stream)).expect("read BIFF records");
    for rec in records.iter().filter(|r| r.record_type == biff::records::FORMULA) {
        assert!(rec.data.len() >= 22, "FORMULA record too short");
        let rw = u16::from_le_bytes([rec.data[0], rec.data[1]]);
        let cl = u16::from_le_bytes([rec.data[2], rec.data[3]]);
        if rw == row && cl == col {
            let cce = u16::from_le_bytes([rec.data[20], rec.data[21]]) as usize;
            assert!(rec.data.len() >= 22 + cce, "token stream truncated");
            return rec.data[22..22 + cce].to_vec();
        }
    }
    panic!("no FORMULA record at row {row} col {col}");
}

#[test]
fn named_constant_formula_uses_value_class_name_ptg() {
    let mut wb = Workbook::new();
    wb.define_name("MyTax", "0.07").unwrap();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=MyTax*2").unwrap();
    ws.set_formula_result(0, 0, CellValue::Number(0.14))
        .unwrap();

    let tokens = only_formula_tokens(&wb);
    assert_eq!(
        tokens.first().copied(),
        Some(0x43),
        "scalar defined names must emit V-class PtgName (0x43); tokens={tokens:02X?}"
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
        Some(0x29),
        "SUM over a reference expression must start with PtgMemFunc; tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(3).copied(),
        Some(0x23),
        "left NameRef must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(8).copied(),
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
        Some(0x29),
        "SUM over a reference expression must start with PtgMemFunc; tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(3).copied(),
        Some(0x23),
        "left NameRef must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(8).copied(),
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
        Some(0x29),
        "SUM over a reference expression must start with PtgMemFunc; tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(3).copied(),
        Some(0x23),
        "left endpoint must emit R-class PtgName (0x23); tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(8).copied(),
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
fn union_parens_survive_outside_sum() {
    // The parens around a union are semantic, not decorative: they make
    // the union a single argument. =COUNT((A1,B1)) has ONE argument;
    // dropping the parens turns it into a two-argument call, and for
    // multi-area INDEX it shifts every later argument. The paren must
    // survive in any context, not just the SUM PtgMemFunc path.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).unwrap();
    ws.set_cell_value("B1", 2.0).unwrap();
    ws.set_cell_formula("D1", "=COUNT((A1,B1))").unwrap();
    ws.set_formula_result(0, 3, CellValue::Number(2.0)).unwrap();
    ws.set_cell_formula("D2", "=(A1,B1)").unwrap();
    ws.set_formula_result(1, 3, CellValue::Number(1.0)).unwrap();

    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    assert_eq!(
        s.get_formula_at(0, 3).expect("COUNT formula must survive"),
        "=COUNT((A1,B1))",
    );
    assert_eq!(
        s.get_formula_at(1, 3).expect("bare union formula must survive"),
        "=(A1,B1)",
    );
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

#[test]
fn if_3arg_formula_emits_attr_if_skip_chain() {
    // IF(cond, t, f) compiles to an MS-XLS optimized 3-token chain
    // (MS-XLS §2.5.198.39 PtgAttrIf + §2.5.198.37 PtgAttrGoto) that
    // short-circuits one branch. Excel emits:
    //   cond
    //   PtgAttrIf [offset = t_size + 4]
    //   t_branch
    //   PtgAttrSkip [offset = f_size + 7]
    //   f_branch
    //   PtgAttrSkip [offset = 3]
    //   PtgFuncVar(IF, argc=3)
    //
    // For `=IF(A1>0,1,2)`: t = PtgInt(1) = 3 bytes, f = PtgInt(2) = 3.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 5.0).unwrap();
    ws.set_cell_formula("B1", "=IF(A1>0,1,2)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(1.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    // 5 (PtgRef) + 3 (PtgInt) + 1 (PtgGT) + 4 (PtgAttrIf) + 3 (PtgInt)
    //   + 4 (PtgAttrSkip) + 3 (PtgInt) + 4 (PtgAttrSkip) + 4 (PtgFuncVar)
    //   = 31 bytes
    assert_eq!(
        tokens.len(),
        31,
        "expected 31 token bytes for 3-arg IF; got {tokens:02X?}"
    );
    // PtgAttrIf at byte 9, offset 7 = t_size(3) + 4
    assert_eq!(&tokens[9..13], &[0x19, 0x02, 0x07, 0x00]);
    // PtgAttrSkip at byte 16, offset 10 = f_size(3) + 7
    assert_eq!(&tokens[16..20], &[0x19, 0x08, 0x0A, 0x00]);
    // PtgAttrSkip at byte 23, offset 3 (trailing)
    assert_eq!(&tokens[23..27], &[0x19, 0x08, 0x03, 0x00]);
    // PtgFuncVar(IF=1, argc=3) V-class
    assert_eq!(&tokens[27..31], &[0x42, 0x03, 0x01, 0x00]);
}

#[test]
fn if_2arg_formula_emits_attr_if_skip_chain() {
    // 2-arg IF(cond, t): no false branch. PtgAttrIf jumps directly to
    // the trailing PtgAttrSkip + PtgFuncVar(IF, argc=2) if cond is false.
    //   cond
    //   PtgAttrIf [offset = t_size + 4]
    //   t_branch
    //   PtgAttrSkip [offset = 3]
    //   PtgFuncVar(IF, argc=2)
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 5.0).unwrap();
    ws.set_cell_formula("B1", "=IF(A1>0,A1*2)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(10.0))
        .unwrap();

    let tokens = only_formula_tokens(&wb);
    // 5 + 3 + 1 + 4 + (5 PtgRef + 3 PtgInt + 1 PtgMul) + 4 + 4 = 30
    assert_eq!(
        tokens.len(),
        30,
        "expected 30 token bytes for 2-arg IF; got {tokens:02X?}"
    );
    // PtgAttrIf at byte 9, offset 13 = t_size(9) + 4
    assert_eq!(&tokens[9..13], &[0x19, 0x02, 0x0D, 0x00]);
    // PtgAttrSkip at byte 22, offset 3
    assert_eq!(&tokens[22..26], &[0x19, 0x08, 0x03, 0x00]);
    // PtgFuncVar(IF=1, argc=2) V-class
    assert_eq!(&tokens[26..30], &[0x42, 0x02, 0x01, 0x00]);
}

#[test]
fn choose_3_branches_emit_attr_choose_with_jump_table() {
    // CHOOSE(selector, c0, c1, c2) compiles to a PtgAttrChoose with a
    // 4-entry u16 jump table (one entry per choice + final exit entry).
    // Jump offsets are measured from the start of the table to the byte
    // position of each choice (or PtgFuncVar for the exit). Each choice
    // is followed by a PtgAttrSkip whose offset lands at the last byte
    // of the formula (matching IF's pattern). MS-XLS §2.5.198.40
    // PtgAttrChoose.
    //
    // For =CHOOSE(A1,10,20,30) with A1 as the selector and PtgInt(N)
    // for each branch (each 3 bytes), the expected layout is exactly
    // what Excel emits.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_formula("B1", "=CHOOSE(A1,10,20,30)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(20.0))
        .unwrap();

    let tokens = only_formula_tokens(&wb);
    // 5 (PtgRef V A1) + 12 (PtgAttrChoose + 4-entry table) + (3+4)*3 - 4 (no
    //   skip after last choice; replaced by trailing 4) + 4 (PtgFuncVar)
    // = 5 + 12 + 9 + 4 + 9 + 4 + 9 + 4 + 4 = 42
    assert_eq!(
        tokens.len(),
        42,
        "expected 42 token bytes for 3-arg CHOOSE; got {tokens:02X?}"
    );
    // Selector A1 (V class)
    assert_eq!(&tokens[0..5], &[0x44, 0x00, 0x00, 0x00, 0xC0]);
    // PtgAttrChoose with nc=3
    assert_eq!(&tokens[5..9], &[0x19, 0x04, 0x03, 0x00]);
    // Jump table: 8, 15, 22, 29 (each as little-endian u16)
    assert_eq!(
        &tokens[9..17],
        &[0x08, 0x00, 0x0F, 0x00, 0x16, 0x00, 0x1D, 0x00]
    );
    // Choice 0 = PtgInt(10) + PtgAttrSkip offset=17
    assert_eq!(&tokens[17..20], &[0x1E, 0x0A, 0x00]);
    assert_eq!(&tokens[20..24], &[0x19, 0x08, 0x11, 0x00]);
    // Choice 1 = PtgInt(20) + PtgAttrSkip offset=10
    assert_eq!(&tokens[24..27], &[0x1E, 0x14, 0x00]);
    assert_eq!(&tokens[27..31], &[0x19, 0x08, 0x0A, 0x00]);
    // Choice 2 = PtgInt(30) + trailing PtgAttrSkip offset=3
    assert_eq!(&tokens[31..34], &[0x1E, 0x1E, 0x00]);
    assert_eq!(&tokens[34..38], &[0x19, 0x08, 0x03, 0x00]);
    // PtgFuncVar argc=4 iftab=100 (CHOOSE)
    assert_eq!(&tokens[38..42], &[0x42, 0x04, 0x64, 0x00]);
}

#[test]
fn choose_naked_ref_branch_uses_r_class() {
    // CHOOSE choices are val_or_ref like IF's t/f args. Naked PtgRef in
    // a choice position must emit R-class (0x24) to preserve the reference;
    // a value expression (A1*2) emits V-class via the BinaryOp forcing rule.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_formula("B1", "=CHOOSE(A1,A1,A1*2)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(2.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    // Selector at bytes 0..5 = PtgRef V (0x44)
    assert_eq!(tokens[0], 0x44);
    // After PtgAttrChoose (bytes 5-14: 4-byte header + 6-byte 3-entry table
    // for nc=2), choice 0 starts at byte 15. Naked A1 → PtgRef R-class
    // (0x24) preserving the reference.
    assert_eq!(tokens[15], 0x24, "tokens={tokens:02X?}");
    // After PtgAttrSkip (bytes 20-23), choice 1 (A1*2) starts at byte 24 —
    // PtgRef V-class (0x44) because the BinaryOp arm forces V children.
    assert_eq!(tokens[24], 0x44, "tokens={tokens:02X?}");
}

#[test]
fn nested_if_in_if_emits_r_class_inner_func() {
    // Excel: a nested IF in the t/f branch of an outer IF emits the inner
    // IF's PtgFuncVar as R-class (0x22) because IF is reference-class and
    // the branch position is val_or_ref. The outer IF stays V-class (0x42).
    // Verified against Excel-authored bytes for =IF(A1>0,IF(A1>10,1,2),3).
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).unwrap();
    ws.set_cell_formula("B1", "=IF(A1>0,IF(A1>10,1,2),3)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(2.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    // The stream contains two PtgFuncVar(iftab=1) tokens. The inner one
    // must be R-class (0x22), the outer (final) one V-class (0x42).
    let if_var_positions: Vec<usize> = tokens
        .windows(4)
        .enumerate()
        .filter(|(_, w)| (w[0] == 0x22 || w[0] == 0x42) && w[1] == 0x03 && w[2] == 0x01 && w[3] == 0x00)
        .map(|(i, _)| i)
        .collect();
    assert_eq!(
        if_var_positions.len(),
        2,
        "expected two IF PtgFuncVar tokens; tokens={tokens:02X?}"
    );
    // Inner IF (earlier in stream) = R-class
    assert_eq!(
        tokens[if_var_positions[0]], 0x22,
        "inner IF must be R-class (0x22); tokens={tokens:02X?}"
    );
    // Outer IF (last token) = V-class
    assert_eq!(
        tokens[if_var_positions[1]], 0x42,
        "outer IF must be V-class (0x42); tokens={tokens:02X?}"
    );
}

#[test]
fn if_in_value_function_emits_v_class_inner_func() {
    // Excel: =ABS(IF(A1>0,A1,A2)) — ABS's arg is V-class, so the inner IF
    // is emitted V-class (0x42). Contrast with SUM(IF(...)) where the IF is
    // R-class. Verified against Excel-authored bytes.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).unwrap();
    ws.set_cell_value("A2", -3.0).unwrap();
    ws.set_cell_formula("B1", "=ABS(IF(A1>0,A1,A2))").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(4.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    // Inner IF PtgFuncVar must be V-class (0x42).
    let if_var = tokens
        .windows(4)
        .find(|w| (w[0] == 0x22 || w[0] == 0x42) && w[1] == 0x03 && w[2] == 0x01 && w[3] == 0x00)
        .map(|w| w[0]);
    assert_eq!(
        if_var,
        Some(0x42),
        "IF inside value-fn ABS must be V-class; tokens={tokens:02X?}"
    );
    // ABS is fixed-arity V-class PtgFunc (0x41) iftab=24.
    assert_eq!(&tokens[tokens.len() - 3..], &[0x41, 0x18, 0x00]);
}

/// Full FORMULA record token region (rgce + rgcb) for a single-cell formula,
/// for verifying array-constant encoding where data spills into the rgcb.
fn formula_rgce_rgcb(wb: &Workbook) -> (Vec<u8>, Vec<u8>) {
    let bytes = XlsWriter::write_to_bytes(wb).expect("write");
    let cfb = CompoundFile::open(Cursor::new(bytes)).expect("open CFB");
    let stream_path = if cfb.exists("/Workbook") {
        "/Workbook"
    } else {
        "/Book"
    };
    let stream = cfb.read_stream(stream_path).expect("read workbook stream");
    let records = biff::read_all_records(&mut Cursor::new(stream)).expect("read BIFF records");
    let rec = records
        .iter()
        .find(|r| r.record_type == biff::records::FORMULA)
        .expect("a FORMULA record");
    let cce = u16::from_le_bytes([rec.data[20], rec.data[21]]) as usize;
    (
        rec.data[22..22 + cce].to_vec(),
        rec.data[22 + cce..].to_vec(),
    )
}

#[test]
fn array_constant_in_sum_emits_ptg_array_and_rgcb() {
    // =SUM({1,2,3}) → rgce: PtgArray(0x60) + 7 reserved, then PtgAttrSum.
    // rgcb: ncols-1(1B)=2, nrows-1(2B)=0, then three number elements
    // (0x01 + f64). Verified against native Excel authoring.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=SUM({1,2,3})").unwrap();
    ws.set_formula_result(0, 0, CellValue::Number(6.0)).unwrap();

    let (rgce, rgcb) = formula_rgce_rgcb(&wb);
    // rgce: 0x60 + 7 reserved (8 bytes) + PtgAttrSum (0x19 0x10 .. ..) = 12.
    assert_eq!(rgce.len(), 12, "rgce={rgce:02X?}");
    assert_eq!(rgce[0], 0x60, "PtgArray A-class; rgce={rgce:02X?}");
    assert_eq!(&rgce[8..10], &[0x19, 0x10], "PtgAttrSum; rgce={rgce:02X?}");
    // rgcb header: cols-1=2, rows-1=0.
    assert_eq!(&rgcb[0..3], &[0x02, 0x00, 0x00], "rgcb header; rgcb={rgcb:02X?}");
    // three numbers: 0x01 + f64
    assert_eq!(rgcb[3], 0x01);
    assert_eq!(f64::from_le_bytes(rgcb[4..12].try_into().unwrap()), 1.0);
    assert_eq!(rgcb[12], 0x01);
    assert_eq!(f64::from_le_bytes(rgcb[13..21].try_into().unwrap()), 2.0);
    assert_eq!(rgcb[21], 0x01);
    assert_eq!(f64::from_le_bytes(rgcb[22..30].try_into().unwrap()), 3.0);
    assert_eq!(rgcb.len(), 30, "rgcb={rgcb:02X?}");
}

#[test]
fn array_constant_round_trips_text() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=SUM({1,2;3,4})").unwrap();
    ws.set_formula_result(0, 0, CellValue::Number(10.0)).unwrap();

    let parsed = write_then_read(&wb);
    let f = parsed
        .worksheet(0)
        .unwrap()
        .get_formula_at(0, 0)
        .expect("array formula must survive");
    assert!(f.contains("{1,2;3,4}"), "array text lost: {f:?}");
}

#[test]
fn unary_plus_emits_ptg_uplus() {
    // =+A1 → PtgRef, PtgUplus (0x12). Excel preserves the leading plus.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_formula("B1", "=+A1").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(2.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    assert_eq!(tokens[0], 0x44, "operand A1 V-class; {tokens:02X?}");
    assert_eq!(tokens.last().copied(), Some(0x12), "trailing PtgUplus; {tokens:02X?}");
    assert_eq!(tokens.len(), 6);
}

#[test]
fn parentheses_emit_ptg_paren() {
    // =(A1+A2)*2 → A1, A2, Add, PtgParen(0x15), Int(2), Mul. Excel keeps the
    // paren as a postfix PtgParen token.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_value("A2", 3.0).unwrap();
    ws.set_cell_formula("B1", "=(A1+A2)*2").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(10.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    // 5(ref) + 5(ref) + 1(add) + 1(paren) + 3(int) + 1(mul) = 16
    assert_eq!(tokens.len(), 16, "{tokens:02X?}");
    assert_eq!(tokens[10], 0x03, "PtgAdd; {tokens:02X?}");
    assert_eq!(tokens[11], 0x15, "PtgParen; {tokens:02X?}");
    assert_eq!(tokens[15], 0x05, "PtgMul; {tokens:02X?}");
}

#[test]
fn nested_parentheses_emit_multiple_ptg_paren() {
    // =((A1)) → PtgRef, PtgParen, PtgParen. Every paren pair preserved.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_formula("B1", "=((A1))").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(2.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    assert_eq!(tokens.len(), 7, "{tokens:02X?}");
    assert_eq!(&tokens[5..7], &[0x15, 0x15], "two PtgParen; {tokens:02X?}");
}

#[test]
fn vlookup_table_array_emits_r_class() {
    // VLOOKUP(lookup_value, table_array, col_index): the table_array (arg 1)
    // is a reference, emitted R-class PtgArea (0x25); lookup_value (arg 0)
    // and col_index (arg 2) are V-class. Shared FunctionDef fix verified
    // byte-for-byte on XLSB; this pins the XLS side at token level.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_value("A2", 3.0).unwrap();
    ws.set_cell_value("A3", 4.0).unwrap();
    ws.set_cell_formula("B1", "=VLOOKUP(A1,A1:A3,1)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(2.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    // lookup_value A1 → V-class PtgRef (0x44).
    assert_eq!(tokens[0], 0x44, "lookup_value must be V-class; {tokens:02X?}");
    // table_array A1:A3 → R-class PtgArea (0x25) at byte 5.
    assert_eq!(
        tokens[5], 0x25,
        "table_array must be R-class PtgArea; {tokens:02X?}"
    );
    // Ends with PtgFuncVar (0x42) argc=3 iftab=102 (VLOOKUP).
    assert_eq!(&tokens[tokens.len() - 4..], &[0x42, 0x03, 0x66, 0x00]);
}

#[test]
fn index_emits_r_class_array_arg() {
    // INDEX(array, row, [col]): the first arg is a reference, emitted
    // R-class PtgArea (0x25). At top level the INDEX token is V-class
    // (0x42); inside SUM it's R-class (0x22). Verified against native
    // Excel authoring (NOT the resave, which repaired our old V-class
    // arg into a UDF wrapper).
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).unwrap();
    ws.set_cell_value("A2", 5.0).unwrap();
    ws.set_cell_value("A3", 6.0).unwrap();
    ws.set_cell_formula("B1", "=INDEX(A1:A3,1)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(4.0)).unwrap();
    ws.set_cell_formula("B2", "=SUM(INDEX(A1:A3,1))").unwrap();
    ws.set_formula_result(1, 1, CellValue::Number(4.0)).unwrap();

    let top = only_formula_tokens_at(&wb, 0, 1);
    // First token: PtgArea R-class (0x25) for the array arg.
    assert_eq!(top[0], 0x25, "INDEX arg0 must be R-class PtgArea; {top:02X?}");
    // Ends with PtgFuncVar V-class (0x42) argc=2 iftab=29.
    assert_eq!(&top[top.len() - 4..], &[0x42, 0x02, 0x1D, 0x00]);

    let in_sum = only_formula_tokens_at(&wb, 1, 1);
    assert_eq!(in_sum[0], 0x25, "INDEX arg0 R-class in SUM too; {in_sum:02X?}");
    // INDEX token R-class (0x22) inside SUM, then PtgAttrSum.
    assert_eq!(&in_sum[in_sum.len() - 8..], &[0x22, 0x02, 0x1D, 0x00, 0x19, 0x10, 0x00, 0x00]);
}

#[test]
fn sum_of_if_emits_r_class_inner_func() {
    // =SUM(IF(A1>0,A1,A2)) — SUM's arg is R-class and IF is reference-class,
    // so the inner IF is R-class (0x22). Verified against Excel-authored bytes.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).unwrap();
    ws.set_cell_value("A2", -3.0).unwrap();
    ws.set_cell_formula("B1", "=SUM(IF(A1>0,A1,A2))").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(4.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    let if_var = tokens
        .windows(4)
        .find(|w| (w[0] == 0x22 || w[0] == 0x42) && w[1] == 0x03 && w[2] == 0x01 && w[3] == 0x00)
        .map(|w| w[0]);
    assert_eq!(
        if_var,
        Some(0x22),
        "IF inside SUM must be R-class; tokens={tokens:02X?}"
    );
}

#[test]
fn giant_if_branch_falls_back_without_panicking() {
    // End-to-end smoke test: a giant IF must not panic on write/read and
    // must preserve its cached value. NOTE: this does NOT verify the
    // scratch-first no-duplication invariant — any overflow-sized stream is
    // rejected wholesale by the u16 cce limit and falls back to the cached
    // value regardless of whether the optimizer duplicated tokens. That
    // invariant is verified directly by the writer unit tests
    // `emit_optimized_{if,choose}_overflow_leaves_out_untouched`.
    //
    // The branch is a SUM with 22000 flat integer args; this trips the
    // 255-arg PtgFuncVar limit so emit_optimized_if returns Err (not the
    // Ok(false) overflow path), exercising the Err-fallback to a cached cell.
    let mut branch = String::from("SUM(1");
    for _ in 0..22000 {
        branch.push_str(",1");
    }
    branch.push(')');
    let formula = format!("=IF(A1>0,{branch},2)");

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 5.0).unwrap();
    ws.set_cell_formula("B1", &formula).unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(22001.0))
        .unwrap();

    // Must not panic on write or read.
    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    // Cell survives as its cached value (formula too large for a FORMULA
    // record). Either a formula or the cached number is acceptable; what
    // matters is the value is intact and nothing is corrupted.
    let v = s.get_value_at(0, 1);
    assert_eq!(
        v.effective_value(),
        &CellValue::Number(22001.0),
        "giant IF must round-trip its cached value without corruption"
    );
}

#[test]
fn if_formula_text_survives_round_trip() {
    // Beyond byte parity, the formula text must survive a write → read
    // cycle. The decompiler treats PtgAttrIf / PtgAttrSkip / PtgAttrVolatile
    // as no-op hints (they don't push values on the stack), so the
    // round-tripped formula should be the natural IF call.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 5.0).unwrap();
    ws.set_cell_formula("B1", "=IF(A1>0,1,2)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(1.0)).unwrap();
    ws.set_cell_formula("B2", "=IF(A1<0,A1)").unwrap();
    ws.set_formula_result(1, 1, CellValue::Number(0.0)).unwrap();

    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    assert_eq!(s.get_formula_at(0, 1).as_deref(), Some("=IF(A1>0,1,2)"));
    assert_eq!(s.get_formula_at(1, 1).as_deref(), Some("=IF(A1<0,A1)"));
}

#[test]
fn choose_formula_text_survives_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_formula("B1", "=CHOOSE(A1,10,20,30)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(20.0)).unwrap();

    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    assert_eq!(
        s.get_formula_at(0, 1).as_deref(),
        Some("=CHOOSE(A1,10,20,30)")
    );
}

#[test]
fn date_function_integer_args_emit_ptg_int() {
    // DATE(year, month, day) takes three integer arguments. Year values
    // through 2079 and small month/day integers all fit in u16, so they
    // must emit as PtgInt (0x1E, 3 bytes) rather than PtgNum (0x1F, 9
    // bytes). Excel encodes this way so we match for byte parity and
    // because PtgInt is more compact.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=DATE(2020,1,1)").unwrap();
    ws.set_formula_result(0, 0, CellValue::Number(43831.0))
        .unwrap();

    let tokens = only_formula_tokens(&wb);
    // PtgInt(2020) at byte 0: 0x1E 0xE4 0x07 (2020 = 0x07E4)
    assert_eq!(&tokens[0..3], &[0x1E, 0xE4, 0x07]);
    // PtgInt(1) at byte 3
    assert_eq!(&tokens[3..6], &[0x1E, 0x01, 0x00]);
    // PtgInt(1) at byte 6
    assert_eq!(&tokens[6..9], &[0x1E, 0x01, 0x00]);
    // PtgFunc V-class iftab=65 (DATE) — DATE has fixed_arity=true in the
    // FunctionDef registry, so Excel emits the 3-byte PtgFunc instead of
    // PtgFuncVar.
    assert_eq!(&tokens[9..12], &[0x41, 0x41, 0x00]);
    assert_eq!(tokens.len(), 12);
}

#[test]
fn randbetween_formula_emits_volatile_attr_prefix() {
    // RANDBETWEEN(bottom, top) is volatile per Excel docs — its output
    // depends on the recalculation cycle, not just its operands, so Excel
    // emits PtgAttrVolatile (0x19 0x01 0x00 0x00) as the first token of
    // every formula whose AST calls it. This test pins that emission.
    //
    // RED test for the bug surfaced during the function-metadata audit:
    // the runtime `FunctionRegistry` marks RANDBETWEEN volatile, but the
    // writer's iftab-based `function_is_volatile` list did not, so we
    // omitted the prefix. After the registry unification this passes.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=RANDBETWEEN(1,10)").unwrap();
    ws.set_formula_result(0, 0, CellValue::Number(5.0)).unwrap();

    let tokens = only_formula_tokens(&wb);
    assert_eq!(
        tokens.first().copied(),
        Some(0x19),
        "RANDBETWEEN formula must start with PtgAttrVolatile (0x19); tokens={tokens:02X?}"
    );
    assert_eq!(
        tokens.get(1).copied(),
        Some(0x01),
        "PtgAttrVolatile subtype byte must be 0x01; tokens={tokens:02X?}"
    );
}

#[test]
fn atp_function_emits_namex_funcvar_udf() {
    // Analysis-ToolPak functions (Ftab 384..=476) are not native BIFF8
    // functions: Excel serializes them as an add-in UDF call. The token
    // stream for `=EDATE(A1,12)` is:
    //   PtgNameX(ixti=0, nameindex=1)   39 00 00 01 00 00 00   (7 bytes)
    //   PtgRef A1 (R-class)             24 + 4-byte payload    (5 bytes)
    //   PtgInt(12)                      1E 0C 00               (3 bytes)
    //   PtgFuncVar(argc=3, iftab=0xFF)  42 03 FF 00            (4 bytes)
    // argc = nargs + 1 (the PtgNameX counts as the first operand); iftab
    // 0x00FF is the "user-defined function" sentinel.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 40000.0).unwrap();
    ws.set_cell_formula("B1", "=EDATE(A1,12)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(40000.0)).unwrap();

    let tokens = only_formula_tokens_at(&wb, 0, 1);
    assert_eq!(
        &tokens[0..7],
        &[0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00],
        "expected PtgNameX(ixti=0, nameindex=1, reserved=0); got {tokens:02X?}"
    );
    assert_eq!(
        tokens[7], 0x24,
        "A1 arg must be R-class PtgRef (0x24); got {tokens:02X?}"
    );
    assert_eq!(
        &tokens[12..15],
        &[0x1E, 0x0C, 0x00],
        "expected PtgInt(12); got {tokens:02X?}"
    );
    assert_eq!(
        &tokens[15..19],
        &[0x42, 0x03, 0xFF, 0x00],
        "expected PtgFuncVar(argc=3, iftab=0x00FF); got {tokens:02X?}"
    );
    assert_eq!(tokens.len(), 19, "unexpected token length: {tokens:02X?}");
}

#[test]
fn atp_functions_get_alphabetical_nameindex() {
    // EXTERNNAME records are emitted one per distinct add-in function,
    // sorted alphabetically, with a 1-based nameindex. With EDATE, GCD,
    // and NETWORKDAYS in use the order is EDATE(1), GCD(2), NETWORKDAYS(3).
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 40000.0).unwrap();
    ws.set_cell_value("B1", 41000.0).unwrap();
    ws.set_cell_formula("C1", "=NETWORKDAYS(A1,B1)").unwrap();
    ws.set_cell_formula("C2", "=EDATE(A1,12)").unwrap();
    ws.set_cell_formula("C3", "=GCD(A1,B1)").unwrap();
    ws.set_formula_result(0, 2, CellValue::Number(1.0)).unwrap();
    ws.set_formula_result(1, 2, CellValue::Number(1.0)).unwrap();
    ws.set_formula_result(2, 2, CellValue::Number(1.0)).unwrap();

    let nameindex_of = |row: u16| -> u16 {
        let t = only_formula_tokens_at(&wb, row, 2);
        assert_eq!(t[0], 0x39, "row {row} must start with PtgNameX; got {t:02X?}");
        u16::from_le_bytes([t[3], t[4]])
    };
    assert_eq!(nameindex_of(1), 1, "EDATE should be nameindex 1");
    assert_eq!(nameindex_of(2), 2, "GCD should be nameindex 2");
    assert_eq!(nameindex_of(0), 3, "NETWORKDAYS should be nameindex 3");
}

#[test]
fn atp_function_round_trips_text_via_externname() {
    // After the writer emits a PtgNameX + EXTERNNAME, the reader must
    // resolve the name back through the AddIn SUPBOOK's external-name
    // table (not the defined-name table) so the formula text survives.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 40000.0).unwrap();
    ws.set_cell_formula("B1", "=EDATE(A1,12)").unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(40000.0)).unwrap();

    let parsed = write_then_read(&wb);
    let s = parsed.worksheet(0).unwrap();
    let f = s
        .get_formula_at(0, 1)
        .expect("ATP formula text must survive XLS round-trip");
    assert!(
        f.contains("EDATE") && f.contains("A1"),
        "ATP function EDATE lost from formula: {f:?}"
    );
}
