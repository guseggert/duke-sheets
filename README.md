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
sheet.setCellStyle('A1', {
  font: { bold: true, color: { hex: 'FFFFFF' } },
  fill: { fillType: 'solid', color: { hex: '1F4E79' } },
});

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

50+ accessors for styles, comments, hyperlinks, tables,
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
sheet.set_cell_style("A1", {
    "font": {"bold": True, "color": {"hex": "FFFFFF"}},
    "fill": {"fill_type": "solid", "color": {"hex": "1F4E79"}},
})

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

Same 50+ accessors as the Node.js API: cell styles, formatted values,
comments, hyperlinks, tables, freeze panes, page setup, etc.

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

> Last updated: 2026-07-13 &middot; commit [`188c268`](../../commit/188c268)
>
> `cargo bench --features full -p duke-sheets`

| Group | Case | Library | Time |
|-------|------|---------|------|
| xlsx_read/100_cells | - | calamine | 128.8 µs |
| xlsx_read/100_cells | - | duke-sheets | 245.3 µs |
| xlsx_read/100_cells | - | umya-spreadsheet | 366.7 µs |
| xlsx_read/10k_cells | - | calamine | 7.593 ms |
| xlsx_read/10k_cells | - | duke-sheets | 12.36 ms |
| xlsx_read/10k_cells | - | umya-spreadsheet | 17.33 ms |
| xlsx_read/1k_cells | - | calamine | 804.8 µs |
| xlsx_read/1k_cells | - | duke-sheets | 1.383 ms |
| xlsx_read/1k_cells | - | umya-spreadsheet | 1.98 ms |
| xlsx_write_serialize/100_cells | - | duke-sheets | 516.7 µs |
| xlsx_write_serialize/100_cells | - | umya-spreadsheet | 610.9 µs |
| xlsx_write_serialize/10k_cells | - | duke-sheets | 25.73 ms |
| xlsx_write_serialize/10k_cells | - | umya-spreadsheet | 21.17 ms |
| xlsx_write_serialize/1k_cells | - | duke-sheets | 2.537 ms |
| xlsx_write_serialize/1k_cells | - | umya-spreadsheet | 2.35 ms |
| xlsx_write_full/100_cells | - | duke-sheets | 550 µs |
| xlsx_write_full/100_cells | - | rust_xlsxwriter | 567.3 µs |
| xlsx_write_full/100_cells | - | umya-spreadsheet | 786.2 µs |
| xlsx_write_full/10k_cells | - | duke-sheets | 26.57 ms |
| xlsx_write_full/10k_cells | - | rust_xlsxwriter | 17.43 ms |
| xlsx_write_full/10k_cells | - | umya-spreadsheet | 28.19 ms |
| xlsx_write_full/1k_cells | - | duke-sheets | 2.67 ms |
| xlsx_write_full/1k_cells | - | rust_xlsxwriter | 1.952 ms |
| xlsx_write_full/1k_cells | - | umya-spreadsheet | 3.08 ms |
| csv_read/100_cells | - | duke-sheets | 40.18 µs |
| csv_read/10k_cells | - | duke-sheets | 1.206 ms |
| csv_read/1k_cells | - | duke-sheets | 154.1 µs |
| csv_write/100_cells | - | duke-sheets | 10.78 µs |
| csv_write/10k_cells | - | duke-sheets | 1.189 ms |
| csv_write/1k_cells | - | duke-sheets | 111.7 µs |
| formula_parse/complex | - | - | 19.4 µs |
| formula_parse/medium | - | - | 9.855 µs |
| formula_parse/simple | - | - | 2.999 µs |
| formula_parse/throughput_1000 | - | - | 1.06 ms |
| calculation/linear_chain | 100 | - | 94.34 µs |
| calculation/linear_chain | 500 | - | 469.5 µs |
| calculation/linear_chain | 1000 | - | 933.1 µs |
| calculation/fan_out | 26 | - | 86.49 µs |
| calculation/fan_out | 52 | - | 189.6 µs |
| calculation/fan_out | 100 | - | 419.2 µs |
| calculation/fan_out | 200 | - | 1.057 ms |
| calculation/cross_sheet | 100 | - | 130.5 µs |
| calculation/cross_sheet | 500 | - | 590.9 µs |
| calculation/cross_sheet | 1000 | - | 1.166 ms |
| calculation/cross_sheet | 5000 | - | 4.943 ms |
| calculation/mixed | 100 | - | 195.6 µs |
| calculation/mixed | 500 | - | 841.2 µs |
| calculation/mixed | 1000 | - | 1.681 ms |
| calculation/repeated_lookups | - | repeated_lookups | 310.1 ms |
<!-- BENCHMARKS:END -->

## License

MIT
