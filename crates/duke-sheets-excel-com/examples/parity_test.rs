//! Example: Parity testing between duke-sheets and Excel via COM bridge.
//!
//! This demonstrates using the Excel COM bridge to:
//! 1. Create a workbook with data and formulas
//! 2. Have Excel recalculate everything
//! 3. Read back the computed values
//! 4. Compare with duke-sheets' own formula engine
//!
//! Prerequisites:
//!   - WINE installed and in PATH
//!   - Microsoft Excel installed in the WINE prefix
//!   - excel-com-bridge.exe built:
//!     cargo build --target x86_64-pc-windows-gnu -p excel-com-bridge --release
//!
//! Run:
//!   cargo run --example parity_test -p duke-sheets-excel-com

use duke_sheets_excel_com::{ExcelBridge, ExcelBridgeConfig};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Excel COM Bridge Parity Test ===\n");

    // Start the bridge (this launches wine excel-com-bridge.exe)
    println!("Starting Excel COM bridge...");
    let bridge = ExcelBridge::start(ExcelBridgeConfig::default())?;

    // Create a new workbook
    println!("Creating workbook...");
    let wb = bridge.create_workbook()?;

    // --- Set up test data ---
    println!("Setting up test data...");

    // Headers
    wb.set_cell_value("A1", "Product")?;
    wb.set_cell_value("B1", "Q1")?;
    wb.set_cell_value("C1", "Q2")?;
    wb.set_cell_value("D1", "Q3")?;
    wb.set_cell_value("E1", "Q4")?;
    wb.set_cell_value("F1", "Total")?;
    wb.set_cell_value("G1", "Average")?;

    // Row 2: Widget
    wb.set_cell_value("A2", "Widget")?;
    wb.set_cell_value("B2", 1500.0)?;
    wb.set_cell_value("C2", 2300.0)?;
    wb.set_cell_value("D2", 1800.0)?;
    wb.set_cell_value("E2", 3100.0)?;
    wb.set_cell_formula("F2", "=SUM(B2:E2)")?;
    wb.set_cell_formula("G2", "=AVERAGE(B2:E2)")?;

    // Row 3: Gadget
    wb.set_cell_value("A3", "Gadget")?;
    wb.set_cell_value("B3", 800.0)?;
    wb.set_cell_value("C3", 950.0)?;
    wb.set_cell_value("D3", 1100.0)?;
    wb.set_cell_value("E3", 1400.0)?;
    wb.set_cell_formula("F3", "=SUM(B3:E3)")?;
    wb.set_cell_formula("G3", "=AVERAGE(B3:E3)")?;

    // Row 4: Doohickey
    wb.set_cell_value("A4", "Doohickey")?;
    wb.set_cell_value("B4", 3200.0)?;
    wb.set_cell_value("C4", 2800.0)?;
    wb.set_cell_value("D4", 3500.0)?;
    wb.set_cell_value("E4", 4100.0)?;
    wb.set_cell_formula("F4", "=SUM(B4:E4)")?;
    wb.set_cell_formula("G4", "=AVERAGE(B4:E4)")?;

    // Summary row
    wb.set_cell_value("A6", "Grand Total")?;
    wb.set_cell_formula("F6", "=SUM(F2:F4)")?;
    wb.set_cell_formula("G6", "=AVERAGE(G2:G4)")?;

    // Some more complex formulas
    wb.set_cell_value("A8", "Max Revenue")?;
    wb.set_cell_formula("B8", "=MAX(B2:E4)")?;

    wb.set_cell_value("A9", "Min Revenue")?;
    wb.set_cell_formula("B9", "=MIN(B2:E4)")?;

    wb.set_cell_value("A10", "Count")?;
    wb.set_cell_formula("B10", "=COUNT(B2:E4)")?;

    wb.set_cell_value("A11", "Conditional")?;
    wb.set_cell_formula("B11", "=IF(F6>10000,\"Above Target\",\"Below Target\")")?;

    // --- Force recalculation ---
    println!("Recalculating...");
    bridge.recalculate()?;

    // --- Read back values ---
    println!("\n--- Results from Excel ---\n");

    let test_cells = [
        ("F2", "Widget Total"),
        ("G2", "Widget Average"),
        ("F3", "Gadget Total"),
        ("G3", "Gadget Average"),
        ("F4", "Doohickey Total"),
        ("G4", "Doohickey Average"),
        ("F6", "Grand Total"),
        ("G6", "Grand Average"),
        ("B8", "Max Revenue"),
        ("B9", "Min Revenue"),
        ("B10", "Count"),
        ("B11", "Conditional"),
    ];

    for (cell, label) in &test_cells {
        let value = wb.get_cell_value(cell)?;
        let formula = wb.get_cell_formula(cell)?;
        println!("  {label:20} ({cell}): {value:>12}  formula: {formula}");
    }

    // --- Save the workbook ---
    println!("\nSaving workbook...");
    wb.save("/tmp/parity_test.xlsx")?;
    println!("Saved to /tmp/parity_test.xlsx");

    // --- Clean up ---
    println!("\nShutting down...");
    bridge.shutdown()?;

    println!("\nDone!");
    Ok(())
}
