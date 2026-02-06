"""
duke_sheets - High-performance Excel file library for Python

This package provides fast, memory-efficient access to Excel files (.xlsx)
and CSV files, with full formula calculation support.

Example:
    >>> import duke_sheets
    >>> wb = duke_sheets.Workbook()
    >>> sheet = wb.get_sheet(0)
    >>> sheet.set_cell("A1", 10)
    >>> sheet.set_cell("A2", 20)
    >>> sheet.set_formula("A3", "=A1+A2")
    >>> wb.calculate()
    >>> value = sheet.get_calculated_value("A3")
    >>> print(value.as_number())
    30.0

Classes:
    Workbook: A workbook containing one or more worksheets
    Worksheet: A worksheet within a workbook
    CellValue: Represents a cell value (number, text, boolean, error, formula)
    CalculationStats: Statistics from calculating a workbook
"""

from duke_sheets._native import (
    Workbook,
    Worksheet,
    CellValue,
    CalculationStats,
)

__all__ = [
    "Workbook",
    "Worksheet",
    "CellValue",
    "CalculationStats",
]

__version__ = "0.1.0"
