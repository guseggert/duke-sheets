# Duke Sheets

A Rust library for reading, writing, and manipulating Excel spreadsheets with full formula evaluation.

Includes bindings for:

- Python
- NodeJS
- WebAssembly

## Features

- **File Formats**: `.xlsx`, `.xlsm`, `.xltx`, `.xltm` (Excel 2007+), `.xls` (legacy), `.csv`
- **Formula Engine**: Full evaluation of ~506 Excel functions
- **Charts**: Create, read, and modify chart types
- **Styling**: Fonts, colors, borders, number formats
- **High Performance**: Multi-threaded calculation engine, works in seconds on workbooks with millions of formulas

## Node.js / TypeScript

Install from [npm](https://www.npmjs.com/package/@dukelib/sheets):

```bash
npm install @dukelib/sheets
```
```typescript
import { Workbook } from '@dukelib/sheets';

const wb = new Workbook();
const sheet = wb.getSheet(0);

sheet.setCell('A1', 10);
sheet.setCell('A2', 20);
sheet.setFormula('A3', '=A1+A2');

wb.calculate();
console.log(sheet.getCalculatedValue('A3').asNumber()); // 30

wb.save('output.xlsx');
```

Open existing files from disk, bytes, or CSV strings:

```typescript
const wb = Workbook.open('input.xlsx');
const wb2 = Workbook.fromBytes(buffer);
const wb3 = Workbook.fromCsvString('a,b,c\n1,2,3');
```

Async versions run on separate threads so the event loop stays free:

```typescript
import { openAsync } from '@dukelib/sheets';

const wb = await openAsync('large-file.xlsx');
await wb.calculateAsync();
await wb.saveAsync('output.xlsx');
```

50+ read-only accessors for styles, comments, hyperlinks, tables,
conditional formatting, data validations, merged regions, page setup, and more.

## Python

I haven't publisehd to PyPI yet. You can install from a [GitHub release](https://github.com/guseggert/duke-sheets/releases):

```bash
pip install https://github.com/guseggert/duke-sheets/releases/download/python-v0.1.1/duke_sheets-0.1.1-cp38-abi3-manylinux_2_17_x86_64.manylinux2014_x86_64.whl
```

Wheels are available for Linux (x86_64, aarch64), macOS (x86_64, ARM), and Windows (x64).
Pick the one matching your platform from the release assets.

Or use as an inline script dependency with [uv](https://docs.astral.sh/uv/):

```python
# /// script
# requires-python = ">=3.9"
# dependencies = [
#     "duke-sheets @ https://github.com/guseggert/duke-sheets/releases/download/python-v0.1.1/duke_sheets-0.1.1-cp38-abi3-manylinux_2_17_x86_64.manylinux2014_x86_64.whl",
# ]
# ///

```python
import duke_sheets

wb = duke_sheets.Workbook()
sheet = wb.get_sheet(0)

sheet.set_cell("A1", 10)
sheet.set_cell("A2", 20)
sheet.set_formula("A3", "=A1+A2")

wb.calculate()
print(sheet.get_calculated_value("A3").as_number())  # 30.0

wb.save("output.xlsx")
```

Open existing files:

```python
wb = duke_sheets.Workbook.open("input.xlsx")
wb = duke_sheets.Workbook.from_bytes(data)
wb = duke_sheets.Workbook.from_csv_string("a,b,c\n1,2,3")
```

Same 50+ read-only accessors as the Node.js API: cell styles, formatted
values, comments, hyperlinks, tables, freeze panes, page setup, etc.

## WebAssembly

Install from [npm](https://www.npmjs.com/package/@dukelib/sheets-wasm):

```bash
npm install @dukelib/sheets-wasm
```

```javascript
import { Workbook } from '@dukelib/sheets-wasm';

const wb = new Workbook();
const sheet = wb.getSheet(0);

sheet.setCell('A1', 10);
sheet.setCell('A2', 20);
sheet.setFormula('A3', '=A1+A2');

const stats = wb.calculate();
console.log(sheet.getCalculatedValue('A3').asNumber()); // 30
```

Load files from bytes or CSV:

```javascript
const wb = Workbook.fromBytes(uint8Array);
const wb2 = Workbook.loadCsvString('a,b,c\n1,2,3');

// Export back out
const xlsxBytes = wb.saveXlsxBytes();   // Uint8Array
const csvString = wb.saveCsvString();    // string
```

Full API parity with the Node.js bindings, including all read-only
accessors (returned as plain JS objects via structured serialization).

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

> Last updated: 2026-03-12 &middot; commit [`9617646`](../../commit/9617646)
>
> `cargo bench --features full -p duke-sheets`

| Group | Case | Library | Time |
|-------|------|---------|------|
| xlsx_read/100_cells | — | calamine | 103.6 µs |
| xlsx_read/100_cells | — | duke-sheets | 196.3 µs |
| xlsx_read/100_cells | — | umya-spreadsheet | 343.3 µs |
| xlsx_read/10k_cells | — | calamine | 6.141 ms |
| xlsx_read/10k_cells | — | duke-sheets | 10.73 ms |
| xlsx_read/10k_cells | — | umya-spreadsheet | 14.95 ms |
| xlsx_read/1k_cells | — | calamine | 655.1 µs |
| xlsx_read/1k_cells | — | duke-sheets | 1.169 ms |
| xlsx_read/1k_cells | — | umya-spreadsheet | 1.656 ms |
| xlsx_write_serialize/100_cells | — | duke-sheets | 377.8 µs |
| xlsx_write_serialize/100_cells | — | umya-spreadsheet | 466.4 µs |
| xlsx_write_serialize/10k_cells | — | duke-sheets | 17.56 ms |
| xlsx_write_serialize/10k_cells | — | umya-spreadsheet | 15.55 ms |
| xlsx_write_serialize/1k_cells | — | duke-sheets | 1.779 ms |
| xlsx_write_serialize/1k_cells | — | umya-spreadsheet | 1.723 ms |
| xlsx_write_full/100_cells | — | duke-sheets | 398.9 µs |
| xlsx_write_full/100_cells | — | rust_xlsxwriter | 432.2 µs |
| xlsx_write_full/100_cells | — | umya-spreadsheet | 597.1 µs |
| xlsx_write_full/10k_cells | — | duke-sheets | 18.69 ms |
| xlsx_write_full/10k_cells | — | rust_xlsxwriter | 12.01 ms |
| xlsx_write_full/10k_cells | — | umya-spreadsheet | 21.88 ms |
| xlsx_write_full/1k_cells | — | duke-sheets | 1.912 ms |
| xlsx_write_full/1k_cells | — | rust_xlsxwriter | 1.364 ms |
| xlsx_write_full/1k_cells | — | umya-spreadsheet | 2.367 ms |
| csv_read/100_cells | — | duke-sheets | 28.04 µs |
| csv_read/10k_cells | — | duke-sheets | 1.753 ms |
| csv_read/1k_cells | — | duke-sheets | 166 µs |
| csv_write/100_cells | — | duke-sheets | 9.521 µs |
| csv_write/10k_cells | — | duke-sheets | 1.301 ms |
| csv_write/1k_cells | — | duke-sheets | 118.4 µs |
| formula_parse/complex | — | — | 12.22 µs |
| formula_parse/medium | — | — | 6.172 µs |
| formula_parse/simple | — | — | 1.657 µs |
| formula_parse/throughput_1000 | — | — | 647.8 µs |
| calculation/linear_chain | 100 | — | 91.78 µs |
| calculation/linear_chain | 500 | — | 510.4 µs |
| calculation/linear_chain | 1000 | — | 1.037 ms |
| calculation/fan_out | 26 | — | 39.41 µs |
| calculation/fan_out | 52 | — | 87.81 µs |
| calculation/fan_out | 100 | — | 206.5 µs |
| calculation/fan_out | 200 | — | 569.8 µs |
| calculation/cross_sheet | 100 | — | 128.8 µs |
| calculation/cross_sheet | 500 | — | 642 µs |
| calculation/cross_sheet | 1000 | — | 1.287 ms |
| calculation/cross_sheet | 5000 | — | 6.151 ms |
| calculation/mixed | 100 | — | 192.5 µs |
| calculation/mixed | 500 | — | 957.8 µs |
| calculation/mixed | 1000 | — | 1.919 ms |
<!-- BENCHMARKS:END -->

## License

MIT OR Apache-2.0
