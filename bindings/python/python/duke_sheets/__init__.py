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
    CellValue: Represents a cell value (number, text, boolean, error, or cached formula result)
    CalculationStats: Statistics from calculating a workbook
"""

from duke_sheets._native import (
    Alignment,
    BorderEdge,
    BorderStyle,
    CalculationStats,
    CellProtection,
    CellValue,
    Chart,
    ChartEx,
    ChildTransform,
    Color,
    Drawing,
    DrawingAnchor,
    DrawingComment,
    DrawingGroup,
    DrawingMeta,
    DrawingText,
    EmbeddedImage,
    FillStyle,
    FontStyle,
    FormControl,
    FormControlInteractionResult,
    GradientStop,
    GroupTransform,
    NumberFormat,
    ProtectedRange,
    RawDrawing,
    RawDrawingRelationship,
    RichTextRun,
    RunFont,
    SheetProtection,
    Shape,
    ShapeFill,
    ShapeLine,
    Style,
    Workbook,
    WorkbookProtection,
    Worksheet,
)
from duke_sheets.types import (
    AlignmentInput,
    BorderEdgeInput,
    BorderStyleInput,
    CellProtectionInput,
    ColorInput,
    FillStyleInput,
    FontStyleInput,
    GradientStopInput,
    NumberFormatInput,
    ProtectedRangeInput,
    SheetProtectionInput,
    StyleInput,
    WorkbookProtectionInput,
)

__all__ = [
    "Alignment",
    "AlignmentInput",
    "BorderEdge",
    "BorderEdgeInput",
    "BorderStyle",
    "BorderStyleInput",
    "CalculationStats",
    "CellProtection",
    "CellProtectionInput",
    "CellValue",
    "Chart",
    "ChartEx",
    "ChildTransform",
    "Color",
    "ColorInput",
    "Drawing",
    "DrawingAnchor",
    "DrawingComment",
    "DrawingGroup",
    "DrawingMeta",
    "DrawingText",
    "EmbeddedImage",
    "FillStyle",
    "FillStyleInput",
    "FontStyle",
    "FontStyleInput",
    "FormControl",
    "FormControlInteractionResult",
    "GradientStop",
    "GradientStopInput",
    "GroupTransform",
    "NumberFormat",
    "NumberFormatInput",
    "ProtectedRange",
    "ProtectedRangeInput",
    "RawDrawing",
    "RawDrawingRelationship",
    "RichTextRun",
    "RunFont",
    "SheetProtection",
    "SheetProtectionInput",
    "Shape",
    "ShapeFill",
    "ShapeLine",
    "Style",
    "StyleInput",
    "Workbook",
    "WorkbookProtection",
    "WorkbookProtectionInput",
    "Worksheet",
]

__version__ = "0.1.0"
