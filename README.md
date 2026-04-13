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

Install from [PyPI](https://pypi.org/project/duke-sheets/):

```bash
pip install duke-sheets
```

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

> Last updated: 2026-04-06 &middot; commit [`fa7c76b`](../../commit/fa7c76b)
>
> `cargo bench --features full -p duke-sheets`

| Group | Case | Library | Time |
|-------|------|---------|------|
| xlsx_read/100_cells | — | calamine | 106.6 µs |
| xlsx_read/100_cells | — | duke-sheets | 215.5 µs |
| xlsx_read/100_cells | — | umya-spreadsheet | 333.6 µs |
| xlsx_read/10k_cells | — | calamine | 6.276 ms |
| xlsx_read/10k_cells | — | duke-sheets | 10.55 ms |
| xlsx_read/10k_cells | — | umya-spreadsheet | 15.19 ms |
| xlsx_read/1k_cells | — | calamine | 664.2 µs |
| xlsx_read/1k_cells | — | duke-sheets | 1.171 ms |
| xlsx_read/1k_cells | — | umya-spreadsheet | 1.694 ms |
| xlsx_write_serialize/100_cells | — | duke-sheets | 388.3 µs |
| xlsx_write_serialize/100_cells | — | umya-spreadsheet | 465.3 µs |
| xlsx_write_serialize/10k_cells | — | duke-sheets | 20.68 ms |
| xlsx_write_serialize/10k_cells | — | umya-spreadsheet | 17.98 ms |
| xlsx_write_serialize/1k_cells | — | duke-sheets | 1.985 ms |
| xlsx_write_serialize/1k_cells | — | umya-spreadsheet | 1.892 ms |
| xlsx_write_full/100_cells | — | duke-sheets | 406.3 µs |
| xlsx_write_full/100_cells | — | rust_xlsxwriter | 456.7 µs |
| xlsx_write_full/100_cells | — | umya-spreadsheet | 583.7 µs |
| xlsx_write_full/10k_cells | — | duke-sheets | 21.48 ms |
| xlsx_write_full/10k_cells | — | rust_xlsxwriter | 14.11 ms |
| xlsx_write_full/10k_cells | — | umya-spreadsheet | 23.81 ms |
| xlsx_write_full/1k_cells | — | duke-sheets | 2.084 ms |
| xlsx_write_full/1k_cells | — | rust_xlsxwriter | 1.503 ms |
| xlsx_write_full/1k_cells | — | umya-spreadsheet | 2.491 ms |
| csv_read/100_cells | — | duke-sheets | 33.23 µs |
| csv_read/10k_cells | — | duke-sheets | 1.097 ms |
| csv_read/1k_cells | — | duke-sheets | 139.8 µs |
| csv_write/100_cells | — | duke-sheets | 9.551 µs |
| csv_write/10k_cells | — | duke-sheets | 1.079 ms |
| csv_write/1k_cells | — | duke-sheets | 96.97 µs |
| formula_parse/complex | — | — | 12.81 µs |
| formula_parse/medium | — | — | 6.525 µs |
| formula_parse/simple | — | — | 1.8 µs |
| formula_parse/throughput_1000 | — | — | 681.1 µs |
| calculation/linear_chain | 100 | — | 79.11 µs |
| calculation/linear_chain | 500 | — | 398.7 µs |
| calculation/linear_chain | 1000 | — | 803.6 µs |
| calculation/fan_out | 26 | — | 67.41 µs |
| calculation/fan_out | 52 | — | 148.6 µs |
| calculation/fan_out | 100 | — | 330.3 µs |
| calculation/fan_out | 200 | — | 833.5 µs |
| calculation/cross_sheet | 100 | — | 111.8 µs |
| calculation/cross_sheet | 500 | — | 511.3 µs |
| calculation/cross_sheet | 1000 | — | 1.019 ms |
| calculation/cross_sheet | 5000 | — | 4.43 ms |
| calculation/mixed | 100 | — | 163.6 µs |
| calculation/mixed | 500 | — | 720.3 µs |
| calculation/mixed | 1000 | — | 1.43 ms |
| calculation/repeated_lookups | — | repeated_lookups | 269.5 ms |
<!-- BENCHMARKS:END -->

## License

MIT OR Apache-2.0
