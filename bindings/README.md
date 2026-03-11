# duke-sheets Bindings

This directory contains language bindings for duke-sheets:

- **[Python](./python/)** - PyO3-based native Python extension
- **[Node.js](./nodejs/)** - NAPI-RS native Node.js/TypeScript addon
- **[WASM](./wasm/)** - WebAssembly bindings for browsers and Node.js

## Quick Start

### Prerequisites

All tooling is managed via [mise](https://mise.jdx.dev/). Install mise first, then:

```bash
# From the repository root, set up all dependencies
mise run setup
```

This will:
1. Add the `wasm32-unknown-unknown` target to your Rust toolchain
2. Install `maturin` via `uv tool`
3. Set up the Python virtual environment

### Build All Bindings

```bash
mise run build
```

### Test All Bindings

```bash
mise run test
```

### Individual Commands

```bash
# Python
mise run build:python     # Build Python bindings
mise run test:python      # Test Python bindings

# WASM
mise run build:wasm       # Build WASM bindings (dev mode)
mise run build:wasm:release  # Build WASM bindings (optimized)
mise run test:wasm        # Test WASM bindings (browser)
mise run test:wasm:node   # Test WASM bindings (Node.js)

# Cleanup
mise run clean            # Clean all build artifacts
mise run clean:python     # Clean Python artifacts only
mise run clean:wasm       # Clean WASM artifacts only
```

---

## Python Bindings

### Installation (Development)

```bash
cd bindings/python
uv sync                   # Install dependencies
uv run maturin develop    # Build and install in development mode
```

### Usage

```python
import duke_sheets

# Create a new workbook
wb = duke_sheets.Workbook()

# Access the first worksheet
sheet = wb.get_sheet(0)

# Set cell values
sheet.set_cell("A1", 10)
sheet.set_cell("A2", 20)
sheet.set_cell("B1", "Hello")
sheet.set_cell("C1", True)

# Set a formula
sheet.set_formula("A3", "=A1+A2")

# Calculate all formulas
stats = wb.calculate()
print(f"Calculated {stats.cells_calculated} cells")

# Get calculated value
result = sheet.get_calculated_value("A3")
print(f"A3 = {result.as_number()}")  # 30.0

# Save to file
wb.save("output.xlsx")

# Open existing file
wb2 = duke_sheets.Workbook.open("existing.xlsx")
```

### API Reference

#### Workbook

| Method | Description |
|--------|-------------|
| `Workbook()` | Create new empty workbook |
| `Workbook.open(path)` | Open from file (.xlsx, .csv) |
| `save(path)` | Save to file |
| `sheet_count` | Number of worksheets |
| `sheet_names` | List of sheet names |
| `get_sheet(index_or_name)` | Get worksheet |
| `add_sheet(name)` | Add new worksheet |
| `remove_sheet(index)` | Remove worksheet |
| `calculate()` | Calculate all formulas |
| `calculate_with_options(...)` | Calculate with custom settings |
| `define_name(name, refers_to)` | Define named range |
| `get_named_range(name)` | Get named range definition |

#### Worksheet

| Method | Description |
|--------|-------------|
| `name` | Worksheet name |
| `set_cell(address, value)` | Set cell value |
| `set_formula(address, formula)` | Set cell formula |
| `get_cell(address)` | Get raw cell value |
| `get_calculated_value(address)` | Get calculated value |
| `get_formula_at(row, col)` | Get formula text for a cell |
| `used_range` | Tuple of (min_row, min_col, max_row, max_col) |
| `set_row_height(row, height)` | Set row height |
| `set_column_width(col, width)` | Set column width |
| `merge_cells(range)` | Merge cells |
| `unmerge_cells(range)` | Unmerge cells |

#### CellValue

| Property/Method | Description |
|-----------------|-------------|
| `is_empty`, `is_number`, `is_text`, etc. | Type checking |
| `as_number()`, `as_text()`, `as_boolean()` | Type conversion |
| `to_python()` | Convert to Python native type |
| Formula text lives on `Worksheet.get_formula_at(row, col)` | Formula metadata access |


---

## Node.js/TypeScript Bindings

See the [main project README](../README.md) for full Node.js documentation
including async API, read-only accessors, and usage examples.

### Installation (Development)

```bash
cd bindings/nodejs
npm install
npm run build              # Build native addon (release)
npm run build:debug        # Build native addon (debug)
npm test                   # Run tests (119 tests)
```

### Package: `@dukelib/sheets`

- **Async API**: `openAsync()`, `calculateAsync()`, `saveAsync()` — run on libuv thread pool
- **Read-only accessors**: 50+ methods for inspecting styles, comments, hyperlinks, tables, etc.
- **Multi-platform CI**: GitHub Actions builds for Linux, macOS, and Windows
- **Consumable from GitHub Releases** without publishing to npm
---

## WASM Bindings

### Building

```bash
cd bindings/wasm
wasm-pack build --target web     # For browsers
wasm-pack build --target nodejs  # For Node.js
```

### Usage (Browser)

```html
<script type="module">
  import init, { Workbook } from './pkg/duke_sheets_wasm.js';

  async function main() {
    await init();

    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", 10);
    sheet.setCell("A2", 20);
    sheet.setFormula("A3", "=A1+A2");

    wb.calculate();

    const result = sheet.getCalculatedValue("A3");
    console.log(`A3 = ${result.asNumber()}`); // 30
  }

  main();
</script>
```

### Usage (Node.js)

```javascript
const { Workbook } = require('./pkg/duke_sheets_wasm.js');

const wb = new Workbook();
const sheet = wb.getSheet(0);

sheet.setCell("A1", 10);
sheet.setCell("A2", 20);
sheet.setFormula("A3", "=A1+A2");

wb.calculate();

console.log(sheet.getCalculatedValue("A3").asNumber()); // 30
```

### Loading/Saving Files

Since browsers don't have filesystem access, use byte arrays:

```javascript
// Load from Uint8Array
const response = await fetch('spreadsheet.xlsx');
const bytes = new Uint8Array(await response.arrayBuffer());
const wb = Workbook.fromBytes(bytes);

// Save to Uint8Array
const outputBytes = wb.saveXlsxBytes();
// Use Blob to download: new Blob([outputBytes], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})

// CSV support
const wb2 = Workbook.loadCsvString("a,b,c\n1,2,3");
const csvOutput = wb.saveCsvString();
```

### TypeScript

TypeScript definitions are automatically generated in `pkg/duke_sheets_wasm.d.ts`.

### API Reference

#### Workbook

| Method | Description |
|--------|-------------|
| `new Workbook()` | Create new empty workbook |
| `Workbook.fromBytes(data)` | Load from Uint8Array (auto-detects XLSX/XLS) |
| `Workbook.loadCsvString(csv)` | Load from CSV string |
| `saveXlsxBytes()` | Save as Uint8Array |
| `saveCsvString()` | Save as CSV string |
| `sheetCount` | Number of worksheets |
| `sheetNames` | Array of sheet names |
| `getSheet(index)` | Get worksheet by index |
| `getSheetByName(name)` | Get worksheet by name |
| `addSheet(name)` | Add new worksheet |
| `removeSheet(index)` | Remove worksheet |
| `calculate()` | Calculate all formulas |
| `calculateWithOptions(...)` | Calculate with custom settings |
| `defineName(name, refersTo)` | Define named range |
| `getNamedRange(name)` | Get named range definition |

#### Worksheet

| Method | Description |
|--------|-------------|
| `name` | Worksheet name |
| `setCell(address, value)` | Set cell value |
| `setFormula(address, formula)` | Set cell formula |
| `getCell(address)` | Get raw cell value |
| `getCalculatedValue(address)` | Get calculated value |
| `getFormulaAt(row, col)` | Get formula text for a cell |
| `usedRange()` | Array [minRow, minCol, maxRow, maxCol] or null |
| `setRowHeight(row, height)` | Set row height |
| `setColumnWidth(col, width)` | Set column width |
| `mergeCells(range)` | Merge cells |
| `unmergeCells(range)` | Unmerge cells |

#### CellValue

| Property/Method | Description |
|-----------------|-------------|
| `isEmpty`, `isNumber`, `isText`, etc. | Type checking |
| `asNumber()`, `asText()`, `asBoolean()` | Type conversion |
| `toJs()` | Convert to JavaScript value |
| `toString()` | String representation |
| Formula text lives on `Worksheet.getFormulaAt(row, col)` | Formula metadata access |

---

## Development

### Project Structure

```
bindings/
├── README.md           # This file
├── python/
│   ├── Cargo.toml      # Rust crate config
│   ├── pyproject.toml  # Python package config (uv/maturin)
│   ├── src/
│   │   └── lib.rs      # PyO3 bindings
│   ├── python/
│   │   └── duke_sheets/
│   │       ├── __init__.py
│   │       └── py.typed
│   └── tests/
│       ├── conftest.py
│       ├── test_workbook.py
│       ├── test_worksheet.py
│       ├── test_cells.py
│       ├── test_formulas.py
│       └── test_calculation.py
├── nodejs/
│   ├── Cargo.toml      # Rust crate config
│   ├── package.json    # @dukelib/sheets package config
│   ├── build.rs        # napi-build setup
│   ├── src/
│   │   ├── lib.rs      # NAPI-RS bindings + async tasks
│   │   ├── types.rs    # NAPI object types (30+ structs)
│   │   ├── workbook_read.rs  # Read-only Workbook methods
│   │   └── worksheet_read.rs # Read-only Worksheet methods
│   ├── npm/            # Platform-specific packages (5 targets)
│   └── __test__/
│       ├── index.spec.ts    # CRUD tests
│       ├── readonly.spec.ts # Read-only API tests
│       └── async.spec.ts    # Async API tests
└── wasm/
    ├── Cargo.toml      # Rust crate config
    ├── src/
    │   └── lib.rs      # wasm-bindgen bindings
    ├── tests/
    │   └── wasm.rs     # wasm-bindgen-test tests
    └── pkg/            # Generated by wasm-pack

### Running Tests Locally

```bash
# Python tests
cd bindings/python
uv run maturin develop
uv run pytest tests/ -v

# Node.js tests
cd bindings/nodejs
npm install
npm run build
npx vitest run

# WASM tests (Node.js - no browser needed)
cd bindings/wasm
wasm-pack test --node

# WASM tests (browser - requires Firefox)
wasm-pack test --headless --firefox
```

### Debugging

For Python, you can build with debug symbols:

```bash
cd bindings/python
uv run maturin develop --release  # Or without --release for debug build
```

For WASM, enable console logging by adding to Cargo.toml:

```toml
[dependencies]
console_error_panic_hook = "0.1"
```

And in lib.rs:

```rust
#[wasm_bindgen(start)]
pub fn init() {
    console_error_panic_hook::set_once();
}
```
