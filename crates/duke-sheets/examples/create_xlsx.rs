//! Example: Create an xlsx file with formulas

use duke_sheets::prelude::*;
use duke_sheets::WorkbookCalculationExt;

fn main() -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();

    // Add header row
    sheet.set_cell_value("A1", "Name")?;
    sheet.set_cell_value("B1", "Value")?;
    sheet.set_cell_value("C1", "Double")?;

    // Add data rows
    sheet.set_cell_value("A2", "Item 1")?;
    sheet.set_cell_value("B2", 100.0)?;
    sheet.set_cell_formula("C2", "=B2*2")?;

    sheet.set_cell_value("A3", "Item 2")?;
    sheet.set_cell_value("B3", 200.0)?;
    sheet.set_cell_formula("C3", "=B3*2")?;

    // Add total row
    sheet.set_cell_value("A4", "Total")?;
    sheet.set_cell_formula("B4", "=SUM(B2:B3)")?;
    sheet.set_cell_formula("C4", "=SUM(C2:C3)")?;

    // Save the file
    workbook.save("/tmp/test.xlsx")?;
    println!("Created /tmp/test.xlsx");

    // Calculate formulas
    let stats = workbook.calculate()?;
    println!(
        "Calculated {} formulas ({} errors)",
        stats.cells_calculated, stats.errors
    );

    // Show calculated values
    let sheet = workbook.worksheet(0).unwrap();
    println!("\nCalculated values:");
    println!("C2 (=B2*2): {:?}", sheet.get_calculated_value_at(1, 2));
    println!("C3 (=B3*2): {:?}", sheet.get_calculated_value_at(2, 2));
    println!(
        "B4 (=SUM(B2:B3)): {:?}",
        sheet.get_calculated_value_at(3, 1)
    );
    println!(
        "C4 (=SUM(C2:C3)): {:?}",
        sheet.get_calculated_value_at(3, 2)
    );

    Ok(())
}
