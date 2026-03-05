# Duke Sheets

A Rust library for reading, writing, and manipulating Excel spreadsheets with full formula evaluation.

## Features

- **File Formats**: `.xlsx`, `.xlsm`, `.xltx`, `.xltm` (Excel 2007+), `.xls` (legacy), `.csv`
- **Formula Engine**: Full evaluation of ~506 Excel functions
- **Charts**: Create, read, and modify chart types
- **Styling**: Fonts, colors, borders, number formats
- **Large Files**: Streaming APIs for 1M+ cells
- **Bindings**: Node.js/TypeScript, Python, WebAssembly

## Node.js / TypeScript

```bash
npm install @duke-sheets/node
```

```typescript
import { Workbook, openAsync } from '@duke-sheets/node';

// Create a new workbook
const wb = new Workbook();
const sheet = wb.getSheet(0);

sheet.setCell('A1', 10);
sheet.setCell('A2', 20);
sheet.setFormula('A3', '=A1+A2');

wb.calculate();
console.log(sheet.getCalculatedValue('A3').asNumber()); // 30

wb.save('output.xlsx');
```

### Async API

For servers and event-loop-sensitive code, async versions run on the
libuv thread pool and return Promises:

```typescript
import { openAsync } from '@duke-sheets/node';

const wb = await openAsync('large-file.xlsx');
await wb.calculateAsync();
await wb.saveAsync('output.xlsx');
```

### Read-Only Accessors

50+ read-only methods for inspecting workbook contents without modification:

```typescript
const sheet = wb.getSheet(0);

// Cell styles
const style = sheet.getCellStyle('A1');
console.log(style?.font.bold, style?.font.name);

// Formatted values (applies number formats)
sheet.getFormattedValue('B2'); // "1,500.00"

// Comments, hyperlinks, tables
sheet.commentCount;    // number
sheet.hyperlinkCount;  // number
sheet.tables;          // JsTable[]

// Page setup, protection, freeze panes
sheet.pageSetup;       // JsPageSetup
sheet.protection;      // JsSheetProtection | null
sheet.freezePanes;     // JsFreezePanes | null

// Formulas, merged regions, conditional formatting
sheet.formulaCells;         // JsFormulaCell[]
sheet.mergedRegions;        // string[]
sheet.conditionalFormats;   // JsConditionalFormatRule[]
sheet.dataValidations;      // JsDataValidation[]
```

### Opening Files

```typescript
import { Workbook, openAsync } from '@duke-sheets/node';

// Sync
const wb = Workbook.open('input.xlsx');

// From bytes
const wb2 = Workbook.fromXlsxBytes(readFileSync('input.xlsx'));

// From CSV
const wb3 = Workbook.fromCsvString('a,b,c\n1,2,3');
```

Supports `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, and `.csv`.

## Rust

Add to your `Cargo.toml`:

```toml
[dependencies]
duke-sheets = { git = "https://github.com/guseggert/duke-sheets.git", features = ["full"] }
```

```rust
use duke_sheets::prelude::*;

fn main() -> Result<()> {
    let mut workbook = Workbook::new();

    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_name("Sales Data")?;

    sheet.set_cell_value("A1", "Product")?;
    sheet.set_cell_value("B1", "Revenue")?;
    sheet.set_cell_value("A2", "Widget")?;
    sheet.set_cell_value("B2", 1500.0)?;

    sheet.set_cell_formula("B5", "=SUM(B2:B4)")?;

    let header_style = Style::new().bold(true);
    sheet.set_cell_style("A1", &header_style)?;
    sheet.set_cell_style("B1", &header_style)?;

    workbook.save("sales.xlsx")?;

    Ok(())
}
```

## Crate Structure

| Crate | Description |
|-------|-------------|
| `duke-sheets` | Main API crate (re-exports all functionality) |
| `duke-sheets-core` | Core data structures |
| `duke-sheets-formula` | Formula parser and evaluator |
| `duke-sheets-xlsx` | XLSX reader/writer |
| `duke-sheets-xls` | XLS reader (legacy format) |
| `duke-sheets-csv` | CSV reader/writer |
| `duke-sheets-chart` | Chart support |
| `duke-sheets-html` | HTML table export |

<!-- BENCHMARKS:START -->
### Benchmarks

> Last updated: 2026-03-05 &middot; commit [`de043b4`](../../commit/de043b4)
>
> `cargo bench --features full -p duke-sheets`

| Group | Case | Library | Time |
|-------|------|---------|------|
| xlsx_read/100_cells | — | calamine | 81.73 µs |
| xlsx_read/100_cells | — | duke-sheets | 146.8 µs |
| xlsx_read/100_cells | — | umya-spreadsheet | 261.6 µs |
| xlsx_read/10k_cells | — | calamine | 4.329 ms |
| xlsx_read/10k_cells | — | duke-sheets | 7.906 ms |
| xlsx_read/10k_cells | — | umya-spreadsheet | 11.07 ms |
| xlsx_read/1k_cells | — | calamine | 467.4 µs |
| xlsx_read/1k_cells | — | duke-sheets | 859.3 µs |
| xlsx_read/1k_cells | — | umya-spreadsheet | 1.24 ms |
| xlsx_write_serialize/100_cells | — | duke-sheets | 301.9 µs |
| xlsx_write_serialize/100_cells | — | umya-spreadsheet | 377.2 µs |
| xlsx_write_serialize/10k_cells | — | duke-sheets | 13.81 ms |
| xlsx_write_serialize/10k_cells | — | umya-spreadsheet | 12.74 ms |
| xlsx_write_serialize/1k_cells | — | duke-sheets | 1.368 ms |
| xlsx_write_serialize/1k_cells | — | umya-spreadsheet | 1.369 ms |
| xlsx_write_full/100_cells | — | duke-sheets | 311.3 µs |
| xlsx_write_full/100_cells | — | rust_xlsxwriter | 436.2 µs |
| xlsx_write_full/100_cells | — | umya-spreadsheet | 471.8 µs |
| xlsx_write_full/10k_cells | — | duke-sheets | 14.96 ms |
| xlsx_write_full/10k_cells | — | rust_xlsxwriter | 10.76 ms |
| xlsx_write_full/10k_cells | — | umya-spreadsheet | 17.85 ms |
| xlsx_write_full/1k_cells | — | duke-sheets | 1.473 ms |
| xlsx_write_full/1k_cells | — | rust_xlsxwriter | 1.3 ms |
| xlsx_write_full/1k_cells | — | umya-spreadsheet | 1.863 ms |
| csv_read/100_cells | — | duke-sheets | 20.65 µs |
| csv_read/10k_cells | — | duke-sheets | 1.431 ms |
| csv_read/1k_cells | — | duke-sheets | 124.3 µs |
| csv_write/100_cells | — | duke-sheets | 7.416 µs |
| csv_write/10k_cells | — | duke-sheets | 1.089 ms |
| csv_write/1k_cells | — | duke-sheets | 88.81 µs |
| formula_parse/complex | — | — | 9.317 µs |
| formula_parse/medium | — | — | 4.652 µs |
| formula_parse/simple | — | — | 1.276 µs |
| formula_parse/throughput_1000 | — | — | 489.8 µs |
| calculation/linear_chain | 100 | — | 768 µs |
| calculation/linear_chain | 500 | — | 24.17 ms |
| calculation/linear_chain | 1000 | — | 73 ms |
| calculation/fan_out | 26 | — | 115.6 µs |
| calculation/fan_out | 52 | — | 389.8 µs |
| calculation/fan_out | 100 | — | 1.307 ms |
| calculation/fan_out | 200 | — | 4.923 ms |
| calculation/cross_sheet | 100 | — | 178.7 µs |
| calculation/cross_sheet | 500 | — | 925.3 µs |
| calculation/cross_sheet | 1000 | — | 1.908 ms |
| calculation/cross_sheet | 5000 | — | 12.1 ms |
| calculation/mixed | 100 | — | 327.6 µs |
| calculation/mixed | 500 | — | 1.625 ms |
| calculation/mixed | 1000 | — | 3.349 ms |
<!-- BENCHMARKS:END -->

## License

MIT OR Apache-2.0
