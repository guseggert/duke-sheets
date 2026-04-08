use std::collections::HashMap;

use duke_sheets_core::CellAddress;

#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub(super) enum CellFormulaKind {
    #[default]
    Normal,
    Shared,
    Array,
    DataTable,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct CellFormulaState {
    pub(super) kind: CellFormulaKind,
    shared_index: Option<u32>,
    /// The `ref` attribute on array/dataTable formulas (e.g., "A1:A3").
    pub(super) array_ref: Option<String>,
    pub(super) data_table_input1_ref: Option<String>,
    pub(super) data_table_input2_ref: Option<String>,
}

#[derive(Debug, Clone)]
pub(super) struct SharedFormulaMaster {
    pub(super) base_cell_ref: String,
    pub(super) formula: String,
}

pub(super) fn parse_cell_formula_state(e: &quick_xml::events::BytesStart<'_>) -> CellFormulaState {
    let mut state = CellFormulaState::default();

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"t" => {
                if let Ok(v) = attr.unescape_value() {
                    state.kind = match v.as_ref() {
                        "shared" => CellFormulaKind::Shared,
                        "array" => CellFormulaKind::Array,
                        "dataTable" => CellFormulaKind::DataTable,
                        _ => CellFormulaKind::Normal,
                    };
                }
            }
            b"si" => {
                state.shared_index = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<u32>().ok());
            }
            b"r1" => {
                state.data_table_input1_ref = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"r2" => {
                state.data_table_input2_ref = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"ref" => {
                state.array_ref = attr.unescape_value().ok().map(|s| s.to_string());
            }
            _ => {}
        }
    }

    state
}

pub(super) fn resolve_cell_formula(
    cell_ref: &str,
    formula: Option<&str>,
    formula_state: &CellFormulaState,
    shared_formula_masters: &mut HashMap<u32, SharedFormulaMaster>,
) -> Option<String> {
    match formula_state.kind {
        CellFormulaKind::Normal | CellFormulaKind::Array => formula.map(|f| f.to_string()),
        CellFormulaKind::DataTable => match formula {
            Some(f) => Some(f.to_string()),
            None => {
                let arg1 = formula_state.data_table_input1_ref.as_deref().unwrap_or("");
                let arg2 = formula_state.data_table_input2_ref.as_deref().unwrap_or("");
                Some(format!("TABLE({},{})", arg1, arg2))
            }
        },
        CellFormulaKind::Shared => {
            let si = formula_state.shared_index?;
            if let Some(f) = formula {
                shared_formula_masters.insert(
                    si,
                    SharedFormulaMaster {
                        base_cell_ref: cell_ref.to_string(),
                        formula: f.to_string(),
                    },
                );
                Some(f.to_string())
            } else {
                let master = shared_formula_masters.get(&si)?;
                Some(translate_shared_formula(
                    &master.formula,
                    &master.base_cell_ref,
                    cell_ref,
                ))
            }
        }
    }
}

pub(super) fn translate_shared_formula(
    formula: &str,
    base_cell_ref: &str,
    cell_ref: &str,
) -> String {
    let base = match CellAddress::parse(base_cell_ref) {
        Ok(v) => v,
        Err(_) => return formula.to_string(),
    };
    let target = match CellAddress::parse(cell_ref) {
        Ok(v) => v,
        Err(_) => return formula.to_string(),
    };

    let row_delta = target.row as i32 - base.row as i32;
    let col_delta = target.col as i32 - base.col as i32;

    shift_a1_references(formula, row_delta, col_delta)
}

pub(super) fn shift_a1_references(formula: &str, row_delta: i32, col_delta: i32) -> String {
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;
    let mut in_string = false;

    while i < bytes.len() {
        if bytes[i] >= 0x80 {
            let ch = formula[i..].chars().next().unwrap();
            out.push(ch);
            i += ch.len_utf8();
            continue;
        }
        let ch = bytes[i] as char;

        if ch == '"' {
            in_string = !in_string;
            out.push(ch);
            i += 1;
            continue;
        }

        if !in_string {
            if i > 0 {
                let prev = bytes[i - 1] as char;
                if prev.is_ascii_alphanumeric() || prev == '_' || prev == '.' {
                    out.push(ch);
                    i += 1;
                    continue;
                }
            }
            if let Some((consumed, shifted)) =
                try_shift_cell_ref(&formula[i..], row_delta, col_delta)
            {
                out.push_str(&shifted);
                i += consumed;
                continue;
            }
        }

        out.push(ch);
        i += 1;
    }

    out
}

pub(super) fn try_shift_cell_ref(
    s: &str,
    row_delta: i32,
    col_delta: i32,
) -> Option<(usize, String)> {
    let b = s.as_bytes();
    let mut i = 0usize;

    let col_abs = if b.get(i) == Some(&b'$') {
        i += 1;
        true
    } else {
        false
    };

    let col_start = i;
    while let Some(&c) = b.get(i) {
        if (c as char).is_ascii_uppercase() {
            i += 1;
        } else {
            break;
        }
    }
    if i == col_start {
        return None;
    }

    let col_letters = &s[col_start..i];
    let mut col = a1_col_to_index(col_letters)? as i32;

    let row_abs = if b.get(i) == Some(&b'$') {
        i += 1;
        true
    } else {
        false
    };

    let row_start = i;
    while let Some(&c) = b.get(i) {
        if (c as char).is_ascii_digit() {
            i += 1;
        } else {
            break;
        }
    }
    if i == row_start {
        return None;
    }

    let mut row: i32 = s[row_start..i].parse::<i32>().ok()?.saturating_sub(1);

    if let Some(&next) = b.get(i) {
        let next = next as char;
        if next.is_ascii_alphanumeric() || next == '_' || next == '.' {
            return None;
        }
    }

    if !col_abs {
        col += col_delta;
    }
    if !row_abs {
        row += row_delta;
    }

    if col < 0 || row < 0 {
        return Some((i, "#REF!".to_string()));
    }

    let mut shifted = String::new();
    if col_abs {
        shifted.push('$');
    }
    shifted.push_str(&a1_index_to_col(col as u16));
    if row_abs {
        shifted.push('$');
    }
    shifted.push_str(&(row as u32 + 1).to_string());

    Some((i, shifted))
}

pub(super) fn a1_col_to_index(col: &str) -> Option<u16> {
    let mut value: u32 = 0;
    for ch in col.chars() {
        if !ch.is_ascii_uppercase() {
            return None;
        }
        value = value
            .saturating_mul(26)
            .saturating_add((ch as u8 - b'A' + 1) as u32);
    }
    if value == 0 {
        None
    } else {
        u16::try_from(value - 1).ok()
    }
}

pub(super) fn a1_index_to_col(mut index: u16) -> String {
    let mut col = String::new();
    loop {
        let rem = (index % 26) as u8;
        col.insert(0, (b'A' + rem) as char);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    col
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use crate::reader::XlsxReader;
    use duke_sheets_core::{CellAddress, CellValue};

    fn build_single_sheet_xlsx(sheet_xml: &str) -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(sheet_xml.as_bytes()).unwrap();

            zip.finish().unwrap();
        }
        buf
    }

    fn formula_text_at<'a>(
        sheet: &'a duke_sheets_core::Worksheet,
        address: &str,
    ) -> Option<&'a str> {
        let addr = CellAddress::parse(address).unwrap();
        sheet.get_formula_at(addr.row, addr.col)
    }

    fn is_array_formula_at(sheet: &duke_sheets_core::Worksheet, address: &str) -> bool {
        let addr = CellAddress::parse(address).unwrap();
        sheet
            .formula_data_at(addr.row, addr.col)
            .map(|formula| formula.is_array_formula())
            .unwrap_or(false)
    }

    #[test]
    fn test_read_shared_formula_master_and_follower() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="n"><v>1</v></c>
      <c r="B1" t="n"><v>2</v></c>
      <c r="C1"><f t="shared" si="0">A1+B1</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="n"><v>4</v></c>
      <c r="B2" t="n"><v>5</v></c>
      <c r="C2"><f t="shared" si="0"/><v>9</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(formula_text_at(sheet, "C1"), Some("=A1+B1"));
        assert_eq!(formula_text_at(sheet, "C2"), Some("=A2+B2"));
    }

    #[test]
    fn test_read_shared_formula_preserves_absolute_and_shifts_ranges() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="n"><v>1</v></c>
      <c r="B1" t="n"><v>2</v></c>
      <c r="D1"><f t="shared" si="1">SUM($A$1:B1)+LEN("A1")</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="n"><v>4</v></c>
      <c r="B2" t="n"><v>5</v></c>
      <c r="D2"><f t="shared" si="1"/><v>6</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            formula_text_at(sheet, "D1"),
            Some("=SUM($A$1:B1)+LEN(\"A1\")")
        );
        assert_eq!(
            formula_text_at(sheet, "D2"),
            Some("=SUM($A$1:B2)+LEN(\"A1\")")
        );
    }

    #[test]
    fn test_read_array_formula_anchor_and_spill() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:A3">ROW(A1:A3)</f><v>1</v></c>
    </row>
    <row r="2"><c r="A2"><v>2</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Anchor cell: has formula, array_result populated
        let a1 = sheet.get_value("A1").unwrap();
        assert_eq!(formula_text_at(sheet, "A1"), Some("=ROW(A1:A3)"));
        assert!(
            is_array_formula_at(sheet, "A1"),
            "A1 should be an array formula"
        );
        assert_eq!(a1.as_number(), Some(1.0));

        // Non-anchor cells: replicated formula with their own cached values
        let a2 = sheet.get_value("A2").unwrap();
        assert_eq!(
            formula_text_at(sheet, "A2"),
            Some("=ROW(A1:A3)"),
            "A2 should have the array formula"
        );
        assert_eq!(a2.as_number(), Some(2.0));

        let a3 = sheet.get_value("A3").unwrap();
        assert_eq!(
            formula_text_at(sheet, "A3"),
            Some("=ROW(A1:A3)"),
            "A3 should have the array formula"
        );
        assert_eq!(a3.as_number(), Some(3.0));
    }

    #[test]
    fn test_read_array_formula_2d() {
        // 2D array formula spanning A1:B2
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:B2">A1:B2*2</f><v>10</v></c>
      <c r="B1"><v>20</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>30</v></c>
      <c r="B2"><v>40</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Anchor has array_result with all 4 values
        let a1 = sheet.get_value("A1").unwrap();
        assert!(
            is_array_formula_at(sheet, "A1"),
            "A1 should have array_result"
        );
        assert_eq!(a1.as_number(), Some(10.0));

        // All cells have formula + their own cached value
        assert_eq!(formula_text_at(sheet, "B1"), Some("=A1:B2*2"));
        assert_eq!(sheet.get_value("B1").unwrap().as_number(), Some(20.0));
        assert_eq!(formula_text_at(sheet, "A2"), Some("=A1:B2*2"));
        assert_eq!(sheet.get_value("A2").unwrap().as_number(), Some(30.0));
        assert_eq!(formula_text_at(sheet, "B2"), Some("=A1:B2*2"));
        assert_eq!(sheet.get_value("B2").unwrap().as_number(), Some(40.0));
    }

    #[test]
    fn test_read_datatable_formula_ref_range() {
        // DataTable with ref range: anchor + non-anchor cells
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="dataTable" ref="A1:A3" r1="C1"/><v>42</v></c>
    </row>
    <row r="2"><c r="A2"><v>84</v></c></row>
    <row r="3"><c r="A3"><v>126</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Anchor: TABLE formula
        let a1 = sheet.get_value("A1").unwrap();
        assert_eq!(formula_text_at(sheet, "A1"), Some("=TABLE(C1,)"));
        assert_eq!(a1.as_number(), Some(42.0));

        // Non-anchor cells: replicated TABLE formula with their own cached values
        let a2 = sheet.get_value("A2").unwrap();
        assert_eq!(
            formula_text_at(sheet, "A2"),
            Some("=TABLE(C1,)"),
            "A2 should have TABLE formula"
        );
        assert_eq!(a2.as_number(), Some(84.0));

        let a3 = sheet.get_value("A3").unwrap();
        assert_eq!(
            formula_text_at(sheet, "A3"),
            Some("=TABLE(C1,)"),
            "A3 should have TABLE formula"
        );
        assert_eq!(a3.as_number(), Some(126.0));
    }

    #[test]
    fn test_read_datatable_formula_two_inputs() {
        // DataTable with both r1 and r2
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="B2"><f t="dataTable" ref="B2:C3" r1="A1" r2="A2"/><v>10</v></c>
      <c r="C2"><v>20</v></c>
    </row>
    <row r="2">
      <c r="B3"><v>30</v></c>
      <c r="C3"><v>40</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(formula_text_at(sheet, "B2"), Some("=TABLE(A1,A2)"));
        assert_eq!(formula_text_at(sheet, "C2"), Some("=TABLE(A1,A2)"));
        assert_eq!(formula_text_at(sheet, "B3"), Some("=TABLE(A1,A2)"));
        assert_eq!(formula_text_at(sheet, "C3"), Some("=TABLE(A1,A2)"));

        assert_eq!(sheet.get_value("B2").unwrap().as_number(), Some(10.0));
        assert_eq!(sheet.get_value("C2").unwrap().as_number(), Some(20.0));
        assert_eq!(sheet.get_value("B3").unwrap().as_number(), Some(30.0));
        assert_eq!(sheet.get_value("C3").unwrap().as_number(), Some(40.0));
    }

    #[test]
    fn test_read_array_formula_single_cell_ref() {
        // Array formula with ref=single cell — no spill targets needed
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1">SUM(B1:B5)</f><v>15</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        let a1 = sheet.get_value("A1").unwrap();
        assert_eq!(formula_text_at(sheet, "A1"), Some("=SUM(B1:B5)"));
        assert_eq!(a1.as_number(), Some(15.0));
        assert!(is_array_formula_at(sheet, "A1"));
    }

    #[test]
    fn test_read_dynamic_array_column() {
        // Dynamic array: cm="1" on anchor, cm="2" on ghosts (column spill)
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1"><f>SEQUENCE(3)</f><v>1</v></c>
    </row>
    <row r="2"><c r="A2" cm="2"><v>2</v></c></row>
    <row r="3"><c r="A3" cm="2"><v>3</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        let a1 = sheet.get_value("A1").unwrap();
        assert_eq!(formula_text_at(sheet, "A1"), Some("=SEQUENCE(3)"));
        assert!(
            is_array_formula_at(sheet, "A1"),
            "anchor should have array_result"
        );
        assert_eq!(a1.as_number(), Some(1.0));

        let a2 = sheet.get_value("A2").unwrap();
        assert!(a2.is_spill_target(), "A2 should be SpillTarget");
        assert_eq!(a2.spill_source(), Some((0, 0)));

        let a3 = sheet.get_value("A3").unwrap();
        assert!(a3.is_spill_target(), "A3 should be SpillTarget");
        assert_eq!(a3.spill_source(), Some((0, 0)));

        // SpillTarget resolution returns the actual values
        assert_eq!(sheet.get_value_at(1, 0), CellValue::Number(2.0));
        assert_eq!(sheet.get_value_at(2, 0), CellValue::Number(3.0));
    }

    #[test]
    fn test_read_dynamic_array_2d() {
        // 2D dynamic array: 2 rows x 3 cols
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1"><f>SEQUENCE(2,3)</f><v>1</v></c>
      <c r="B1" cm="2"><v>2</v></c>
      <c r="C1" cm="2"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" cm="2"><v>4</v></c>
      <c r="B2" cm="2"><v>5</v></c>
      <c r="C2" cm="2"><v>6</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        let a1 = sheet.get_value("A1").unwrap();
        assert!(
            is_array_formula_at(sheet, "A1"),
            "anchor should have array_result"
        );
        assert_eq!(a1.as_number(), Some(1.0));

        // All non-anchor cells are SpillTargets
        for (cell_ref, expected_source) in [
            ("B1", (0u32, 0u16)),
            ("C1", (0, 0)),
            ("A2", (0, 0)),
            ("B2", (0, 0)),
            ("C2", (0, 0)),
        ] {
            let val = sheet.get_value(cell_ref).unwrap();
            assert!(val.is_spill_target(), "{cell_ref} should be SpillTarget");
            assert_eq!(
                val.spill_source(),
                Some(expected_source),
                "{cell_ref} source"
            );
        }

        // Resolution gives correct values
        assert_eq!(sheet.get_value_at(0, 1), CellValue::Number(2.0));
        assert_eq!(sheet.get_value_at(0, 2), CellValue::Number(3.0));
        assert_eq!(sheet.get_value_at(1, 0), CellValue::Number(4.0));
        assert_eq!(sheet.get_value_at(1, 1), CellValue::Number(5.0));
        assert_eq!(sheet.get_value_at(1, 2), CellValue::Number(6.0));
    }

    #[test]
    fn test_read_dynamic_array_single_cell() {
        // Single-cell dynamic array: cm="1" with no ghosts
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1"><f>TODAY()</f><v>45000</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        let a1 = sheet.get_value("A1").unwrap();
        assert_eq!(formula_text_at(sheet, "A1"), Some("=TODAY()"));
        assert!(
            is_array_formula_at(sheet, "A1"),
            "single-cell anchor should have array_result"
        );
        assert_eq!(a1.as_number(), Some(45000.0));
    }

    #[test]
    fn test_read_dynamic_array_with_strings() {
        // Dynamic array with string values
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1" t="str"><f>UNIQUE(B1:B3)</f><v>apple</v></c>
    </row>
    <row r="2"><c r="A2" cm="2" t="str"><v>banana</v></c></row>
    <row r="3"><c r="A3" cm="2" t="str"><v>cherry</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert!(is_array_formula_at(sheet, "A1"));

        let a2 = sheet.get_value("A2").unwrap();
        assert!(a2.is_spill_target());

        // Resolved values
        assert_eq!(
            sheet.get_value_at(1, 0),
            CellValue::String("banana".to_string().into())
        );
        assert_eq!(
            sheet.get_value_at(2, 0),
            CellValue::String("cherry".to_string().into())
        );
    }
}
