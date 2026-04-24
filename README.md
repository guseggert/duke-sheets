# Duke Sheets

A library for reading, writing, manipulating, and evaluating Excel spreadsheets.

Includes bindings for:

- [Python](https://pypi.org/project/duke-sheets/)
- [NodeJS](https://www.npmjs.com/package/@dukelib/sheets)
- [WebAssembly](https://www.npmjs.com/package/@dukelib/sheets-wasm)
- Rust

Duke Sheets includes an extensive test suite:

- Formula tests covering Excel's documentation cases
- Compatibility & parity tests against both LibreOffice and Excel
- Fuzz testing
- Performance benchmarks
- Corpus testing on real-world spreadsheets

Duke Sheets has a multithreaded formula engine which can evaluate millions of formulas per second, and has been profiled against some of the most complex financial spreadsheets in the world.

Supported file formats: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlsb`, `.xls`, `.csv`

Duke Sheets supports all formulas, except ones that don't make sense such as `CALL` and `REGISTER.ID`. Even formulas such as [WEBSERVICE](https://support.microsoft.com/en-us/office/webservice-function-0546a35a-ecc6-4739-aed7-c0b7ce1562c4) are supported. Most workbook metadata is also supported such as formatting, images, charts, etc. Some advanced features are still in progress (e.g., pivot tables).

> [!WARNING]
> Duke Sheets is in alpha. Its API is not yet stable.

### Feature Coverage

See [FEATURES.md](FEATURES.md) for the per-feature support matrix.

## Node.js / TypeScript

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

<!-- BENCHMARKS:START -->
### Benchmarks

> Last updated: 2026-04-24 &middot; commit [`098a89d`](../../commit/098a89d)
>
> `cargo bench --features full -p duke-sheets`

| Group | Case | Library | Time |
|-------|------|---------|------|
| xlsx_read/100_cells | - | calamine | 111.1 µs |
| xlsx_read/100_cells | - | duke-sheets | 217.5 µs |
| xlsx_read/100_cells | - | umya-spreadsheet | 336.4 µs |
| xlsx_read/10k_cells | - | calamine | 6.316 ms |
| xlsx_read/10k_cells | - | duke-sheets | 10.6 ms |
| xlsx_read/10k_cells | - | umya-spreadsheet | 14.7 ms |
| xlsx_read/1k_cells | - | calamine | 682.4 µs |
| xlsx_read/1k_cells | - | duke-sheets | 1.179 ms |
| xlsx_read/1k_cells | - | umya-spreadsheet | 1.668 ms |
| xlsx_write_serialize/100_cells | - | duke-sheets | 390 µs |
| xlsx_write_serialize/100_cells | - | umya-spreadsheet | 460.6 µs |
| xlsx_write_serialize/10k_cells | - | duke-sheets | 20.65 ms |
| xlsx_write_serialize/10k_cells | - | umya-spreadsheet | 17.55 ms |
| xlsx_write_serialize/1k_cells | - | duke-sheets | 1.953 ms |
| xlsx_write_serialize/1k_cells | - | umya-spreadsheet | 1.859 ms |
| xlsx_write_full/100_cells | - | duke-sheets | 406.8 µs |
| xlsx_write_full/100_cells | - | rust_xlsxwriter | 452.8 µs |
| xlsx_write_full/100_cells | - | umya-spreadsheet | 579.2 µs |
| xlsx_write_full/10k_cells | - | duke-sheets | 21.45 ms |
| xlsx_write_full/10k_cells | - | rust_xlsxwriter | 13.82 ms |
| xlsx_write_full/10k_cells | - | umya-spreadsheet | 23.49 ms |
| xlsx_write_full/1k_cells | - | duke-sheets | 2.067 ms |
| xlsx_write_full/1k_cells | - | rust_xlsxwriter | 1.478 ms |
| xlsx_write_full/1k_cells | - | umya-spreadsheet | 2.461 ms |
| csv_read/100_cells | - | duke-sheets | 32.45 µs |
| csv_read/10k_cells | - | duke-sheets | 1.057 ms |
| csv_read/1k_cells | - | duke-sheets | 135.3 µs |
| csv_write/100_cells | - | duke-sheets | 9.818 µs |
| csv_write/10k_cells | - | duke-sheets | 1.137 ms |
| csv_write/1k_cells | - | duke-sheets | 97.06 µs |
| formula_parse/complex | - | - | 12.85 µs |
| formula_parse/medium | - | - | 6.474 µs |
| formula_parse/simple | - | - | 1.783 µs |
| formula_parse/throughput_1000 | - | - | 680 µs |
| calculation/linear_chain | 100 | - | 78.39 µs |
| calculation/linear_chain | 500 | - | 391.8 µs |
| calculation/linear_chain | 1000 | - | 789.2 µs |
| calculation/fan_out | 26 | - | 66.83 µs |
| calculation/fan_out | 52 | - | 147 µs |
| calculation/fan_out | 100 | - | 326.8 µs |
| calculation/fan_out | 200 | - | 823.5 µs |
| calculation/cross_sheet | 100 | - | 107.9 µs |
| calculation/cross_sheet | 500 | - | 494.3 µs |
| calculation/cross_sheet | 1000 | - | 985.3 µs |
| calculation/cross_sheet | 5000 | - | 4.266 ms |
| calculation/mixed | 100 | - | 161 µs |
| calculation/mixed | 500 | - | 709.5 µs |
| calculation/mixed | 1000 | - | 1.405 ms |
| calculation/repeated_lookups | - | repeated_lookups | 267.7 ms |
<!-- BENCHMARKS:END -->

## License

MIT
