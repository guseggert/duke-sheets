# duke-sheets

High-performance Excel file library for Node.js, powered by Rust.

## Installation

```bash
npm install duke-sheets
```

## Quick Start

```typescript
import { Workbook } from 'duke-sheets';

// Create a new workbook
const wb = new Workbook();
const sheet = wb.getSheet(0);

// Set cell values
sheet.setCell('A1', 10);
sheet.setCell('A2', 20);
sheet.setFormula('A3', '=A1+A2');

// Calculate formulas
wb.calculate();

// Get the result
const result = sheet.getCalculatedValue('A3');
console.log(result.asNumber()); // 30

// Save to file
wb.save('output.xlsx');
```

## Opening Existing Files

```typescript
// Open an Excel file
const wb = Workbook.open('input.xlsx');

// Or from bytes (Buffer)
import { readFileSync } from 'fs';
const bytes = readFileSync('input.xlsx');
const wb2 = Workbook.fromXlsxBytes(bytes);

// Or from CSV string
const wb3 = Workbook.fromCsvString('a,b,c\n1,2,3');
```

## Features

- Read and write Excel files (.xlsx, .xls)
- Read and write CSV files
- Full formula calculation engine (506 Excel functions)
- Support for named ranges
- Cell merging
- Row heights and column widths
- Native performance via Rust + NAPI-RS
- Full TypeScript type definitions (auto-generated)

## License

MIT OR Apache-2.0
