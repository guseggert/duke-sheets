use duke_sheets::{Workbook, Worksheet};

pub fn build_fixture(name: &str) -> Workbook {
    match name {
        "repeated-lookups" => build_repeated_lookup_fixture(4000, 2500),
        _ => panic!("Unknown fixture: {name}"),
    }
}

pub fn build_repeated_lookup_fixture(calc_rows: u32, lookup_rows: u32) -> Workbook {
    let mut wb = Workbook::new();
    let _ = wb.worksheet_mut(0).unwrap().set_name("Calc");
    let _ = wb.add_worksheet_with_name("Lookup");

    {
        let sheet = wb.worksheet_mut(1).unwrap();
        for (col, header) in ["Value A", "Value B", "Lookup Key", "Status"]
            .iter()
            .enumerate()
        {
            let _ = sheet.set_cell_value_at(0, col as u16, *header);
        }
        for row in 1..=lookup_rows {
            let r = row - 1;
            let _ = sheet.set_cell_value_at(row, 0, format!("A-{:05}", r % 500));
            let _ = sheet.set_cell_value_at(row, 1, format!("B-{:05}", r % 500));
            let _ = sheet.set_cell_value_at(row, 2, format!("KEY-{:05}", r));
            let _ = sheet.set_cell_value_at(row, 3, if r % 2 == 0 { "Open" } else { "Closed" });
        }
    }

    {
        let sheet = wb.worksheet_mut(0).unwrap();
        write_calc_headers(sheet);

        let lookup_end = lookup_rows + 1;
        for row in 1..=calc_rows {
            let excel_row = row + 1;
            let lookup_id = (row - 1) % lookup_rows;
            let _ = sheet.set_cell_value_at(row, 0, format!("KEY-{:05}", lookup_id));

            for (col, header_col) in [(1u16, 'B'), (2u16, 'C'), (3u16, 'D')] {
                let formula = format!(
                    "=IFERROR(INDEX(Lookup!$A$2:$D${},MATCH($A{},Lookup!$C$2:$C${},0),MATCH({}$1,Lookup!$A$1:$D$1,0)),\"-\")",
                    lookup_end, excel_row, lookup_end, header_col
                );
                let _ = sheet.set_cell_formula_at(row, col, &formula);
            }

            let _ = sheet.set_cell_formula_at(
                row,
                4,
                &format!(
                    "=IFERROR(_xlfn.XLOOKUP($A{},Lookup!$C$2:$C${},Lookup!$B$2:$B${}),\"-\")",
                    excel_row, lookup_end, lookup_end
                ),
            );
        }
    }

    wb
}

fn write_calc_headers(sheet: &mut Worksheet) {
    for (col, header) in [
        (0u16, "Lookup Key"),
        (1u16, "Value A"),
        (2u16, "Value B"),
        (3u16, "Status"),
        (4u16, "XLookup B"),
    ] {
        let _ = sheet.set_cell_value_at(0, col, header);
    }
}
