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

Install from a [GitHub release](https://github.com/guseggert/duke-sheets/releases):

```bash
npm install https://github.com/guseggert/duke-sheets/releases/download/node-v0.1.0/dukelib-sheets-0.1.0.tgz
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
const wb2 = Workbook.fromXlsxBytes(buffer);
const wb3 = Workbook.fromCsvString('a,b,c\n1,2,3');
```

Async versions run on the libuv thread pool so the event loop stays free:

```typescript
import { openAsync } from '@dukelib/sheets';

const wb = await openAsync('large-file.xlsx');
await wb.calculateAsync();
await wb.saveAsync('output.xlsx');
```

50+ read-only accessors for styles, comments, hyperlinks, tables,
conditional formatting, data validations, merged regions, page setup, and more.

## Python

Install from a [GitHub release](https://github.com/guseggert/duke-sheets/releases):

```bash
pip install https://github.com/guseggert/duke-sheets/releases/download/python-v0.1.0/duke_sheets-0.1.0-cp38-abi3-manylinux_2_17_x86_64.manylinux2014_x86_64.whl
```

Wheels are available for Linux (x86_64, aarch64), macOS (x86_64, ARM), and Windows (x64).
Pick the one matching your platform from the release assets.

Or use as an inline script dependency with [uv](https://docs.astral.sh/uv/):

```python
# /// script
# requires-python = ">=3.9"
# dependencies = [
#     "duke-sheets @ https://github.com/guseggert/duke-sheets/releases/download/python-v0.1.0/duke_sheets-0.1.0-cp38-abi3-manylinux_2_17_x86_64.manylinux2014_x86_64.whl",
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
wb = duke_sheets.Workbook.from_xlsx_bytes(data)
wb = duke_sheets.Workbook.from_csv_string("a,b,c\n1,2,3")
```

Same 50+ read-only accessors as the Node.js API: cell styles, formatted
values, comments, hyperlinks, tables, freeze panes, page setup, etc.

## WebAssembly

Install from a [GitHub release](https://github.com/guseggert/duke-sheets/releases):

```bash
# For webpack/vite (bundler target)
npm install https://github.com/guseggert/duke-sheets/releases/download/wasm-v0.1.0/dukelib-sheets-wasm-bundler.tgz

# For <script type="module"> (web target)
npm install https://github.com/guseggert/duke-sheets/releases/download/wasm-v0.1.0/dukelib-sheets-wasm-web.tgz

# For Node.js via WASM
npm install https://github.com/guseggert/duke-sheets/releases/download/wasm-v0.1.0/dukelib-sheets-wasm-nodejs.tgz
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
const wb = Workbook.fromXlsxBytes(uint8Array);
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

> Last updated: 2026-03-10 &middot; commit [`7c96c30`](../../commit/7c96c30)
>
> `cargo bench --features full -p duke-sheets`

| Group | Case | Library | Time |
|-------|------|---------|------|
| xlsx_read/100_cells | — | calamine | 112.3 µs |
| xlsx_read/100_cells | — | duke-sheets | 194.4 µs |
| xlsx_read/100_cells | — | umya-spreadsheet | 312.2 µs |
| xlsx_read/10k_cells | — | calamine | 6.445 ms |
| xlsx_read/10k_cells | — | duke-sheets | 10.92 ms |
| xlsx_read/10k_cells | — | umya-spreadsheet | 15.17 ms |
| xlsx_read/1k_cells | — | calamine | 690.1 µs |
| xlsx_read/1k_cells | — | duke-sheets | 1.176 ms |
| xlsx_read/1k_cells | — | umya-spreadsheet | 1.681 ms |
| xlsx_write_serialize/100_cells | — | duke-sheets | 429.9 µs |
| xlsx_write_serialize/100_cells | — | umya-spreadsheet | 508.4 µs |
| xlsx_write_serialize/10k_cells | — | duke-sheets | 21.2 ms |
| xlsx_write_serialize/10k_cells | — | umya-spreadsheet | 18.31 ms |
| xlsx_write_serialize/1k_cells | — | duke-sheets | 2.148 ms |
| xlsx_write_serialize/1k_cells | — | umya-spreadsheet | 2.024 ms |
| xlsx_write_full/100_cells | — | duke-sheets | 445.2 µs |
| xlsx_write_full/100_cells | — | rust_xlsxwriter | 495.6 µs |
| xlsx_write_full/100_cells | — | umya-spreadsheet | 655.4 µs |
| xlsx_write_full/10k_cells | — | duke-sheets | 22.56 ms |
| xlsx_write_full/10k_cells | — | rust_xlsxwriter | 15.22 ms |
| xlsx_write_full/10k_cells | — | umya-spreadsheet | 24.75 ms |
| xlsx_write_full/1k_cells | — | duke-sheets | 2.278 ms |
| xlsx_write_full/1k_cells | — | rust_xlsxwriter | 1.702 ms |
| xlsx_write_full/1k_cells | — | umya-spreadsheet | 2.665 ms |
| csv_read/100_cells | — | duke-sheets | 31.28 µs |
| csv_read/10k_cells | — | duke-sheets | 1.63 ms |
| csv_read/1k_cells | — | duke-sheets | 123.9 µs |
| csv_write/100_cells | — | duke-sheets | 9.616 µs |
| csv_write/10k_cells | — | duke-sheets | 1.306 ms |
| csv_write/1k_cells | — | duke-sheets | 100.6 µs |
| formula_parse/complex | — | — | 14.12 µs |
| formula_parse/medium | — | — | 7.066 µs |
| formula_parse/simple | — | — | 1.946 µs |
| formula_parse/throughput_1000 | — | — | 731.8 µs |
| calculation/linear_chain | 100 | — | 232.6 µs |
| calculation/linear_chain | 500 | — | 1.261 ms |
| calculation/linear_chain | 1000 | — | 2.57 ms |
| calculation/fan_out | 26 | — | 50.06 µs |
| calculation/fan_out | 52 | — | 115.3 µs |
| calculation/fan_out | 100 | — | 271.9 µs |
| calculation/fan_out | 200 | — | 748.7 µs |
| calculation/cross_sheet | 100 | — | 251 µs |
| calculation/cross_sheet | 500 | — | 1.328 ms |
| calculation/cross_sheet | 1000 | — | 2.873 ms |
| calculation/cross_sheet | 5000 | — | 25.24 ms |
| calculation/mixed | 100 | — | 236.8 µs |
| calculation/mixed | 500 | — | 1.163 ms |
| calculation/mixed | 1000 | — | 2.335 ms |
<!-- BENCHMARKS:END -->

## License

MIT OR Apache-2.0
