# Duke Sheets

A Rust library for reading, writing, and manipulating spreadsheets, with an API similar to Aspose Cells.

## Features

- **File Formats**: XLSX, XLS (legacy), CSV
- **Formula Engine**: Full evaluation of ~450 Excel functions
- **Charts**: Create, read, and modify chart types
- **Styling**: Fonts, colors, borders, number formats
- **Large Files**: Streaming APIs for 1M+ cells
- **Dual API**: Idiomatic Rust + C FFI

## Quick Start

```rust
use duke_sheets::prelude::*;

fn main() -> Result<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    
    // Get the first worksheet
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_name("Sales Data")?;
    
    // Set cell values
    sheet.set_cell_value("A1", "Product")?;
    sheet.set_cell_value("B1", "Revenue")?;
    sheet.set_cell_value("A2", "Widget")?;
    sheet.set_cell_value("B2", 1500.0)?;
    
    // Set a formula
    sheet.set_cell_formula("B5", "=SUM(B2:B4)")?;
    
    // Style the header
    let header_style = Style::new().bold(true);
    sheet.set_cell_style("A1", &header_style)?;
    sheet.set_cell_style("B1", &header_style)?;
    
    // Save to file
    workbook.save("sales.xlsx")?;
    
    Ok(())
}
```

## Installation

Add to your `Cargo.toml`:

```toml
[dependencies]
duke-sheets = "0.1"
```

## Crate Structure

| Crate | Description |
|-------|-------------|
| `duke-sheets` | Main API crate (re-exports all functionality) |
| `duke-sheets-core` | Core data structures |
| `duke-sheets-formula` | Formula parser and evaluator |
| `duke-sheets-xlsx` | XLSX reader/writer |
| `duke-sheets-xls` | XLS reader/writer (optional) |
| `duke-sheets-csv` | CSV reader/writer |
| `duke-sheets-chart` | Chart support |
| `duke-sheets-ffi` | C FFI bindings |

## License

MIT OR Apache-2.0
