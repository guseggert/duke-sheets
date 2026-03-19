# Duke Sheets

A library for reading, writing, and manipulating Excel spreadsheets with full formula evaluation.

Includes bindings for:

- Python
- NodeJS
- WebAssembly
- Rust

Duke Sheets includes an extensive test suite:

- Formula tests covering Excel's documentation cases
- Compatibility tests against both LibreOffice and Excel
- Fuzz testing
- Performance benchmarks
- [Performance regression workflow](docs/PERF_REGRESSION.md)
- Corpus testing on real-world spreadsheets

Duke Sheets has a high-performance multithreaded formula engine which can evaluate millions of formulas in seconds, and has been profiled against some of the most complex financial spreadsheets in the world.

Supported file formats: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.csv`

Additional supported features:

- Styling (fonts, colors, borders, number formatting)

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

> Last updated: 2026-03-16 &middot; commit [`8678329`](../../commit/8678329)
>
> `cargo bench --features full -p duke-sheets`

| Group | Case | Library | Time |
|-------|------|---------|------|
| xlsx_read/100_cells | — | calamine | 110.7 µs |
| xlsx_read/100_cells | — | duke-sheets | 198.2 µs |
| xlsx_read/100_cells | — | umya-spreadsheet | 311.6 µs |
| xlsx_read/10k_cells | — | calamine | 6.552 ms |
| xlsx_read/10k_cells | — | duke-sheets | 10.7 ms |
| xlsx_read/10k_cells | — | umya-spreadsheet | 14.89 ms |
| xlsx_read/1k_cells | — | calamine | 700.8 µs |
| xlsx_read/1k_cells | — | duke-sheets | 1.18 ms |
| xlsx_read/1k_cells | — | umya-spreadsheet | 1.676 ms |
| xlsx_write_serialize/100_cells | — | duke-sheets | 444.2 µs |
| xlsx_write_serialize/100_cells | — | umya-spreadsheet | 512.9 µs |
| xlsx_write_serialize/10k_cells | — | duke-sheets | 22.15 ms |
| xlsx_write_serialize/10k_cells | — | umya-spreadsheet | 18.21 ms |
| xlsx_write_serialize/1k_cells | — | duke-sheets | 2.194 ms |
| xlsx_write_serialize/1k_cells | — | umya-spreadsheet | 2.029 ms |
| xlsx_write_full/100_cells | — | duke-sheets | 466.7 µs |
| xlsx_write_full/100_cells | — | rust_xlsxwriter | 500.9 µs |
| xlsx_write_full/100_cells | — | umya-spreadsheet | 670.6 µs |
| xlsx_write_full/10k_cells | — | duke-sheets | 22.88 ms |
| xlsx_write_full/10k_cells | — | rust_xlsxwriter | 15.06 ms |
| xlsx_write_full/10k_cells | — | umya-spreadsheet | 24.41 ms |
| xlsx_write_full/1k_cells | — | duke-sheets | 2.383 ms |
| xlsx_write_full/1k_cells | — | rust_xlsxwriter | 1.709 ms |
| xlsx_write_full/1k_cells | — | umya-spreadsheet | 2.678 ms |
| csv_read/100_cells | — | duke-sheets | 35.12 µs |
| csv_read/10k_cells | — | duke-sheets | 1.091 ms |
| csv_read/1k_cells | — | duke-sheets | 137 µs |
| csv_write/100_cells | — | duke-sheets | 9.649 µs |
| csv_write/10k_cells | — | duke-sheets | 1.031 ms |
| csv_write/1k_cells | — | duke-sheets | 97.21 µs |
| formula_parse/complex | — | — | 13.82 µs |
| formula_parse/medium | — | — | 7.145 µs |
| formula_parse/simple | — | — | 2.03 µs |
| formula_parse/throughput_1000 | — | — | 743.4 µs |
| calculation/linear_chain | 100 | — | 97.17 µs |
| calculation/linear_chain | 500 | — | 502.6 µs |
| calculation/linear_chain | 1000 | — | 1.014 ms |
| calculation/fan_out | 26 | — | 42.44 µs |
| calculation/fan_out | 52 | — | 92.38 µs |
| calculation/fan_out | 100 | — | 208.4 µs |
| calculation/fan_out | 200 | — | 548.3 µs |
| calculation/cross_sheet | 100 | — | 123.1 µs |
| calculation/cross_sheet | 500 | — | 597.9 µs |
| calculation/cross_sheet | 1000 | — | 1.192 ms |
| calculation/cross_sheet | 5000 | — | 4.706 ms |
| calculation/mixed | 100 | — | 178.2 µs |
| calculation/mixed | 500 | — | 842.5 µs |
| calculation/mixed | 1000 | — | 1.68 ms |
<!-- BENCHMARKS:END -->

## License

MIT OR Apache-2.0
