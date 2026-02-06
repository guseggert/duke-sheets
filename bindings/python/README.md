# duke-sheets

High-performance Excel file library for Python, powered by Rust.

## Installation

```bash
pip install duke-sheets
```

## Quick Start

```python
import duke_sheets

# Create a new workbook
wb = duke_sheets.Workbook()
sheet = wb.get_sheet(0)

# Set cell values
sheet.set_cell("A1", 10)
sheet.set_cell("A2", 20)
sheet.set_formula("A3", "=A1+A2")

# Calculate formulas
wb.calculate()

# Get the result
result = sheet.get_calculated_value("A3")
print(result.as_number())  # 30.0

# Save to file
wb.save("output.xlsx")
```

## Opening Existing Files

```python
# Open an Excel file
wb = duke_sheets.Workbook.open("input.xlsx")

# Or a CSV file
wb = duke_sheets.Workbook.open("data.csv")
```

## Features

- Read and write Excel files (.xlsx)
- Read and write CSV files
- Full formula calculation engine
- Support for named ranges
- Cell merging
- Row heights and column widths

## License

MIT OR Apache-2.0
