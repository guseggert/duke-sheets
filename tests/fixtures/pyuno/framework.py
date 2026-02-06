#!/usr/bin/env python3
"""
PyUNO Fixture Framework for generating XLSX test files.

This provides a clean abstraction for creating test fixtures using LibreOffice's
UNO API. It handles connection management, document creation, and common patterns.

Usage:
    from framework import fixture, FixtureBuilder, StyleSpec

    @fixture("my_test.xlsx")
    def create_my_test(b: FixtureBuilder):
        b.set_sheet_name("TestSheet")
        b.add_data("A1", [10, 20, 30, 40, 50])
        b.add_formula("B1", "=SUM(A1:A5)")
        b.add_comment("A1", "This is a comment", author="Test Author")
        b.add_merge("C1:D2")
        b.set_cell_style("A1", StyleSpec(bold=True, fill_color=0xFFFF00))
        b.add_conditional_format(
            range="A1:A5",
            condition=("greater_than", 25),
            style={"bold": True, "fill_color": 0xFFFF00}
        )

    # Run all fixtures
    if __name__ == "__main__":
        run_fixtures()
"""

import os
import sys
from typing import Dict, List, Any, Callable, Optional, Tuple, Union
from dataclasses import dataclass
from enum import Enum

# UNO imports
try:
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.sheet.ConditionOperator import GREATER, LESS, EQUAL, BETWEEN, NONE
    from com.sun.star.sheet.ValidationAlertStyle import STOP, WARNING, INFO
    from com.sun.star.sheet.ValidationType import (
        LIST as VT_LIST,
        WHOLE as VT_WHOLE,
        DECIMAL as VT_DECIMAL,
        DATE as VT_DATE,
        TIME as VT_TIME,
        TEXT_LEN as VT_TEXT_LEN,
        CUSTOM as VT_CUSTOM,
    )
    from com.sun.star.sheet.ConditionOperator import (
        BETWEEN as VO_BETWEEN,
        NOT_BETWEEN as VO_NOT_BETWEEN,
        EQUAL as VO_EQUAL,
        NOT_EQUAL as VO_NOT_EQUAL,
        GREATER as VO_GREATER,
        LESS as VO_LESS,
        GREATER_EQUAL as VO_GREATER_EQUAL,
        LESS_EQUAL as VO_LESS_EQUAL,
    )

    HAS_UNO = True
except ImportError:
    HAS_UNO = False
    print("WARNING: UNO not available. Running in stub mode.")


# Registry of fixtures
_FIXTURES: Dict[str, Callable] = {}


def fixture(filename: str):
    """Decorator to register a fixture generator function."""

    def decorator(func: Callable):
        _FIXTURES[filename] = func
        return func

    return decorator


def get_fixtures() -> Dict[str, Callable]:
    """Get all registered fixtures."""
    return _FIXTURES.copy()


@dataclass
class StyleSpec:
    """Specification for a cell style."""

    # Font
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = (
        None  # "single", "double", "singleAccounting", "doubleAccounting"
    )
    strikethrough: bool = False
    font_color: Optional[int] = None  # RGB as int, e.g., 0xFF0000
    font_size: Optional[float] = None
    font_name: Optional[str] = None

    # Fill
    fill_color: Optional[int] = None  # RGB as int
    pattern_type: Optional[str] = None  # "solid", "gray125", etc.
    pattern_fg_color: Optional[int] = None
    pattern_bg_color: Optional[int] = None

    # Alignment
    horizontal: Optional[str] = None  # "left", "center", "right", "justify", "fill"
    vertical: Optional[str] = None  # "top", "center", "bottom"
    wrap_text: bool = False
    shrink_to_fit: bool = False
    rotation: int = 0  # -90 to 90, or 255 for vertical
    indent: int = 0

    # Border
    border_style: Optional[str] = None  # "thin", "medium", "thick", "dashed", etc.
    border_color: Optional[int] = None
    left_border: Optional[Tuple[str, int]] = None  # (style, color)
    right_border: Optional[Tuple[str, int]] = None
    top_border: Optional[Tuple[str, int]] = None
    bottom_border: Optional[Tuple[str, int]] = None

    # Number format
    number_format: Optional[str] = None  # e.g., "0.00%", "#,##0"

    # Protection
    locked: Optional[bool] = None
    hidden: Optional[bool] = None


class FixtureBuilder:
    """
    Builder for creating XLSX test fixtures with LibreOffice.

    Provides methods for:
    - Cell data (values, formulas, errors)
    - Cell styling (font, fill, border, alignment)
    - Comments
    - Merged cells
    - Data validation
    - Conditional formatting
    - Multiple sheets
    """

    # Condition type mapping for conditional formatting
    CONDITION_OPERATORS = {
        "greater_than": GREATER if HAS_UNO else 1,
        "less_than": LESS if HAS_UNO else 2,
        "equal": EQUAL if HAS_UNO else 3,
        "between": BETWEEN if HAS_UNO else 4,
        "none": NONE if HAS_UNO else 0,
    }

    # Horizontal alignment mapping
    HORIZONTAL_ALIGN = {
        "general": 0,
        "left": 1,
        "center": 2,
        "right": 3,
        "fill": 4,
        "justify": 5,
        "center_continuous": 6,
        "distributed": 7,
    }

    # Vertical alignment mapping
    VERTICAL_ALIGN = {
        "standard": 0,
        "top": 1,
        "center": 2,
        "bottom": 3,
        "justify": 4,
        "distributed": 5,
    }

    # Border style mapping
    BORDER_STYLES = {
        "none": 0,
        "thin": 1,
        "medium": 2,
        "thick": 3,
        "dashed": 4,
        "dotted": 5,
        "double": 6,
    }

    # Underline style mapping
    UNDERLINE_STYLES = {
        "none": 0,
        "single": 1,
        "double": 2,
        "singleAccounting": 3,
        "doubleAccounting": 4,
    }

    # Validation type mapping
    VALIDATION_TYPES = {
        "list": VT_LIST if HAS_UNO else 1,
        "whole": VT_WHOLE if HAS_UNO else 2,
        "decimal": VT_DECIMAL if HAS_UNO else 3,
        "date": VT_DATE if HAS_UNO else 4,
        "time": VT_TIME if HAS_UNO else 5,
        "text_length": VT_TEXT_LEN if HAS_UNO else 6,
        "custom": VT_CUSTOM if HAS_UNO else 7,
    }

    # Validation operator mapping
    VALIDATION_OPERATORS = {
        "between": VO_BETWEEN if HAS_UNO else 0,
        "not_between": VO_NOT_BETWEEN if HAS_UNO else 1,
        "equal": VO_EQUAL if HAS_UNO else 2,
        "not_equal": VO_NOT_EQUAL if HAS_UNO else 3,
        "greater_than": VO_GREATER if HAS_UNO else 4,
        "less_than": VO_LESS if HAS_UNO else 5,
        "greater_or_equal": VO_GREATER_EQUAL if HAS_UNO else 6,
        "less_or_equal": VO_LESS_EQUAL if HAS_UNO else 7,
    }

    def __init__(self, filename: str, output_dir: Optional[str] = None):
        self.filename = filename
        self.output_dir = output_dir or os.environ.get("OUTPUT_DIR", "/output")
        self.doc = None
        self.sheet = None
        self.sheet_index = 0
        self.desktop = None
        self.ctx = None
        self._style_counter = 0

    def __enter__(self):
        self._connect()
        self._create_document()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.doc:
            try:
                self._save()
            finally:
                self.doc.close(True)
        return False

    def _connect(self):
        """Connect to LibreOffice."""
        if not HAS_UNO:
            raise RuntimeError("UNO not available")

        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )

        self.ctx = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
        )
        smgr = self.ctx.ServiceManager
        self.desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )

    def _create_document(self):
        """Create a new spreadsheet."""
        self.doc = self.desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, ()
        )
        self.sheet = self.doc.getSheets().getByIndex(0)
        self.sheet_index = 0

    def _save(self):
        """Save document as XLSX."""
        filepath = os.path.join(self.output_dir, self.filename)
        if not filepath.startswith("/"):
            filepath = os.path.abspath(filepath)

        url = f"file://{filepath}"
        props = (
            self._make_prop("FilterName", "Calc MS Excel 2007 XML"),
            self._make_prop("Overwrite", True),
        )
        self.doc.storeToURL(url, props)
        print(f"  Created: {self.filename}")

    def _make_prop(self, name: str, value: Any) -> "PropertyValue":
        """Create a UNO PropertyValue."""
        pv = PropertyValue()
        pv.Name = name
        pv.Value = value
        return pv

    def _parse_cell(self, cell_ref: str) -> Tuple[int, int]:
        """Parse cell reference like 'A1' into (col, row) zero-indexed."""
        col = 0
        i = 0

        # Parse column letters
        while i < len(cell_ref) and cell_ref[i].isalpha():
            col = col * 26 + (ord(cell_ref[i].upper()) - ord("A") + 1)
            i += 1
        col -= 1  # Zero-indexed

        # Parse row number
        row = int(cell_ref[i:]) - 1  # Zero-indexed

        return col, row

    def _get_cell(self, cell_ref: str):
        """Get a cell by reference."""
        col, row = self._parse_cell(cell_ref)
        return self.sheet.getCellByPosition(col, row)

    # =========================================================================
    # Sheet Management
    # =========================================================================

    def set_sheet_name(self, name: str):
        """Set the active sheet's name."""
        self.sheet.setName(name)

    def add_sheet(self, name: str):
        """Add a new sheet and make it active."""
        sheets = self.doc.getSheets()
        sheets.insertNewByName(name, sheets.getCount())
        self.sheet_index = sheets.getCount() - 1
        self.sheet = sheets.getByIndex(self.sheet_index)

    def select_sheet(self, index: int):
        """Select a sheet by index."""
        sheets = self.doc.getSheets()
        if index < 0 or index >= sheets.getCount():
            raise ValueError(f"Sheet index {index} out of range")
        self.sheet_index = index
        self.sheet = sheets.getByIndex(index)

    # =========================================================================
    # Cell Data
    # =========================================================================

    def add_data(self, start_cell: str, data: Union[Any, List, List[List]]):
        """
        Add data to the sheet.

        Args:
            start_cell: Starting cell like "A1"
            data: Single value, list of values (column), or list of lists (rows)
        """
        col, row = self._parse_cell(start_cell)

        # Normalize to list of lists
        if not isinstance(data, list):
            data = [[data]]
        elif data and not isinstance(data[0], (list, tuple)):
            data = [[v] for v in data]  # Single column

        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                cell = self.sheet.getCellByPosition(col + col_idx, row + row_idx)
                if isinstance(value, str):
                    cell.setString(value)
                elif isinstance(value, bool):
                    cell.setValue(1 if value else 0)
                    # Also set as formula to preserve boolean type
                    cell.setFormula("TRUE" if value else "FALSE")
                elif value is None:
                    pass  # Leave empty
                else:
                    cell.setValue(value)

    def add_formula(self, cell_ref: str, formula: str):
        """
        Add a formula to a cell.

        Args:
            cell_ref: Cell reference like "A1"
            formula: Formula string like "=SUM(A1:A5)"
        """
        cell = self._get_cell(cell_ref)
        cell.setFormula(formula)

    def add_error(self, cell_ref: str, error_formula: str):
        """
        Add a formula that produces an error.

        Args:
            cell_ref: Cell reference like "A1"
            error_formula: Formula that produces error, e.g., "=1/0" for #DIV/0!
        """
        cell = self._get_cell(cell_ref)
        cell.setFormula(error_formula)

    # =========================================================================
    # Comments
    # =========================================================================

    def add_comment(self, cell_ref: str, text: str, author: Optional[str] = None):
        """
        Add a comment to a cell.

        Args:
            cell_ref: Cell reference like "A1"
            text: Comment text
            author: Optional author name
        """
        col, row = self._parse_cell(cell_ref)

        # Get annotations container
        annotations = self.sheet.getAnnotations()

        # Create cell address
        cell_addr = uno.createUnoStruct("com.sun.star.table.CellAddress")
        cell_addr.Sheet = self.sheet_index
        cell_addr.Column = col
        cell_addr.Row = row

        # Insert annotation
        annotations.insertNew(cell_addr, text)

        # Set author if provided (LibreOffice doesn't directly support per-comment authors
        # in the same way Excel does, but the annotation object may have an Author property)
        if author:
            try:
                annotation = annotations.getByIndex(annotations.getCount() - 1)
                annotation.setAuthor(author)
            except Exception:
                pass  # Author not supported in this LO version

    # =========================================================================
    # Merged Cells
    # =========================================================================

    def add_merge(self, range_ref: str):
        """
        Merge a range of cells.

        Args:
            range_ref: Range reference like "A1:B2"
        """
        cell_range = self.sheet.getCellRangeByName(range_ref)
        cell_range.merge(True)

    # =========================================================================
    # Data Validation
    # =========================================================================

    def add_data_validation(
        self,
        range_ref: str,
        validation_type: str,
        operator: str = "between",
        formula1: str = "",
        formula2: str = "",
        allow_blank: bool = True,
        show_dropdown: bool = True,
        input_title: Optional[str] = None,
        input_message: Optional[str] = None,
        error_title: Optional[str] = None,
        error_message: Optional[str] = None,
        error_style: str = "stop",  # "stop", "warning", "info"
    ):
        """
        Add data validation to a range.

        Args:
            range_ref: Range reference like "A1:A10"
            validation_type: "list", "whole", "decimal", "date", "time", "text_length", "custom"
            operator: "between", "not_between", "equal", "not_equal", "greater_than", etc.
            formula1: First value/formula (or list items for "list" type, comma-separated)
            formula2: Second value/formula (for "between" operators)
            allow_blank: Allow blank cells
            show_dropdown: Show dropdown arrow (for list validation)
            input_title: Title for input message
            input_message: Input message text
            error_title: Title for error message
            error_message: Error message text
            error_style: "stop", "warning", or "info"
        """
        cell_range = self.sheet.getCellRangeByName(range_ref)

        # Get validation object
        validation = cell_range.getPropertyValue("Validation")

        # Set type
        vtype = self.VALIDATION_TYPES.get(validation_type, VT_WHOLE if HAS_UNO else 2)
        validation.setPropertyValue("Type", vtype)

        # Set operator
        vop = self.VALIDATION_OPERATORS.get(operator, VO_BETWEEN if HAS_UNO else 0)
        validation.setPropertyValue("Operator", vop)

        # Set formulas
        if formula1:
            validation.setPropertyValue("Formula1", formula1)
        if formula2:
            validation.setPropertyValue("Formula2", formula2)

        # Set options
        validation.setPropertyValue("IgnoreBlankCells", allow_blank)
        validation.setPropertyValue("ShowList", show_dropdown)

        # Set input message
        if input_title or input_message:
            validation.setPropertyValue("ShowInputMessage", True)
            if input_title:
                validation.setPropertyValue("InputTitle", input_title)
            if input_message:
                validation.setPropertyValue("InputMessage", input_message)

        # Set error message
        if error_title or error_message:
            validation.setPropertyValue("ShowErrorMessage", True)
            if error_title:
                validation.setPropertyValue("ErrorTitle", error_title)
            if error_message:
                validation.setPropertyValue("ErrorMessage", error_message)

        # Set error style
        error_styles = {
            "stop": STOP if HAS_UNO else 0,
            "warning": WARNING if HAS_UNO else 1,
            "info": INFO if HAS_UNO else 2,
        }
        validation.setPropertyValue(
            "ErrorAlertStyle", error_styles.get(error_style, STOP if HAS_UNO else 0)
        )

        # Apply validation
        cell_range.setPropertyValue("Validation", validation)

    # =========================================================================
    # Cell Styling (Direct)
    # =========================================================================

    def set_cell_style(self, cell_ref: str, style: Union[StyleSpec, dict]):
        """
        Apply styling directly to a cell (not via conditional formatting).

        Args:
            cell_ref: Cell reference like "A1"
            style: StyleSpec or dict of style properties
        """
        if isinstance(style, dict):
            style = StyleSpec(**style)

        cell = self._get_cell(cell_ref)
        self._apply_style_to_cell(cell, style)

    def set_range_style(self, range_ref: str, style: Union[StyleSpec, dict]):
        """
        Apply styling directly to a range of cells.

        Args:
            range_ref: Range reference like "A1:B5"
            style: StyleSpec or dict of style properties
        """
        if isinstance(style, dict):
            style = StyleSpec(**style)

        cell_range = self.sheet.getCellRangeByName(range_ref)
        self._apply_style_to_range(cell_range, style)

    def _apply_style_to_cell(self, cell, spec: StyleSpec):
        """Apply a StyleSpec to a cell."""
        # Font properties
        if spec.bold:
            cell.setPropertyValue("CharWeight", 150)  # BOLD
        if spec.italic:
            cell.setPropertyValue("CharPosture", 2)  # ITALIC
        if spec.underline:
            ul_val = self.UNDERLINE_STYLES.get(spec.underline, 1)
            cell.setPropertyValue("CharUnderline", ul_val)
        if spec.strikethrough:
            cell.setPropertyValue("CharStrikeout", 1)
        if spec.font_color is not None:
            cell.setPropertyValue("CharColor", spec.font_color)
        if spec.font_size is not None:
            cell.setPropertyValue("CharHeight", spec.font_size)
        if spec.font_name is not None:
            cell.setPropertyValue("CharFontName", spec.font_name)

        # Fill properties
        if spec.fill_color is not None:
            cell.setPropertyValue("CellBackColor", spec.fill_color)

        # Alignment properties
        if spec.horizontal is not None:
            h_val = self.HORIZONTAL_ALIGN.get(spec.horizontal, 0)
            cell.setPropertyValue("HoriJustify", h_val)
        if spec.vertical is not None:
            v_val = self.VERTICAL_ALIGN.get(spec.vertical, 0)
            cell.setPropertyValue("VertJustify", v_val)
        if spec.wrap_text:
            cell.setPropertyValue("IsTextWrapped", True)
        if spec.shrink_to_fit:
            cell.setPropertyValue("ShrinkToFit", True)
        if spec.rotation != 0:
            # LibreOffice uses 1/100 degree, XLSX uses degrees
            # 255 is special value for vertical text
            if spec.rotation == 255:
                cell.setPropertyValue("Orientation", 1)  # Stacked
            else:
                cell.setPropertyValue("RotateAngle", spec.rotation * 100)
        if spec.indent > 0:
            cell.setPropertyValue("ParaIndent", spec.indent * 200)

        # Border properties
        if spec.border_style is not None:
            self._apply_all_borders(
                cell, spec.border_style, spec.border_color or 0x000000
            )

        # Individual borders
        if spec.left_border:
            self._apply_border(
                cell, "LeftBorder", spec.left_border[0], spec.left_border[1]
            )
        if spec.right_border:
            self._apply_border(
                cell, "RightBorder", spec.right_border[0], spec.right_border[1]
            )
        if spec.top_border:
            self._apply_border(
                cell, "TopBorder", spec.top_border[0], spec.top_border[1]
            )
        if spec.bottom_border:
            self._apply_border(
                cell, "BottomBorder", spec.bottom_border[0], spec.bottom_border[1]
            )

        # Number format
        if spec.number_format is not None:
            number_formats = self.doc.getNumberFormats()
            locale = uno.createUnoStruct("com.sun.star.lang.Locale")
            fmt_id = number_formats.queryKey(spec.number_format, locale, False)
            if fmt_id == -1:
                fmt_id = number_formats.addNew(spec.number_format, locale)
            cell.setPropertyValue("NumberFormat", fmt_id)

    def _apply_style_to_range(self, cell_range, spec: StyleSpec):
        """Apply a StyleSpec to a range (same logic as cell)."""
        # For simplicity, apply to the range object directly
        # Most properties work the same way
        self._apply_style_to_cell(cell_range, spec)

    def _apply_all_borders(self, target, style: str, color: int):
        """Apply the same border to all sides."""
        for side in ["TopBorder", "BottomBorder", "LeftBorder", "RightBorder"]:
            self._apply_border(target, side, style, color)

    def _apply_border(self, target, side: str, style: str, color: int):
        """Apply a border to one side."""
        border_line = uno.createUnoStruct("com.sun.star.table.BorderLine2")
        border_line.LineStyle = self.BORDER_STYLES.get(style, 1)
        border_line.LineWidth = (
            50 if style == "thin" else (100 if style == "medium" else 150)
        )
        border_line.Color = color
        target.setPropertyValue(side, border_line)

    # =========================================================================
    # Conditional Formatting
    # =========================================================================

    def add_conditional_format(
        self,
        range: str,
        condition: Tuple[str, Any],
        style: Union[StyleSpec, dict],
    ):
        """
        Add a conditional format rule.

        Args:
            range: Cell range like "A1:A10"
            condition: Tuple of (operator, value) like ("greater_than", 50)
            style: StyleSpec or dict of style properties
        """
        if isinstance(style, dict):
            style = StyleSpec(**style)

        # Create the cell style for this CF rule
        style_name = self._create_cf_style(style)

        # Get condition operator
        op_name, formula = condition
        operator = self.CONDITION_OPERATORS.get(op_name, NONE if HAS_UNO else 0)

        # Get the range
        cell_range = self.sheet.getCellRangeByName(range)

        # Get conditional format container
        cf_entries = self.sheet.getPropertyValue("ConditionalFormat")

        # Create condition
        props = (
            self._make_prop("Operator", operator),
            self._make_prop("Formula1", str(formula)),
            self._make_prop("StyleName", style_name),
        )
        cf_entries.addNew(props)
        cell_range.setPropertyValue("ConditionalFormat", cf_entries)

    def _create_cf_style(self, spec: StyleSpec) -> str:
        """Create a cell style for conditional formatting and return its name."""
        self._style_counter += 1
        style_name = f"CF_Style_{self._style_counter}"

        # Get style families
        style_families = self.doc.getStyleFamilies()
        cell_styles = style_families.getByName("CellStyles")

        # Create style
        style = self.doc.createInstance("com.sun.star.style.CellStyle")
        cell_styles.insertByName(style_name, style)

        # Apply properties (reuse the same logic)
        self._apply_style_to_cell(style, spec)

        return style_name

    # =========================================================================
    # Row/Column Dimensions (for future use)
    # =========================================================================

    def set_row_height(self, row: int, height: float):
        """
        Set the height of a row.

        Args:
            row: Row number (1-indexed)
            height: Height in points
        """
        rows = self.sheet.getRows()
        row_obj = rows.getByIndex(row - 1)
        # Height is in 1/100 mm, convert from points (1 pt = 0.3528 mm)
        row_obj.setPropertyValue("Height", int(height * 35.28))

    def set_column_width(self, col: Union[int, str], width: float):
        """
        Set the width of a column.

        Args:
            col: Column number (1-indexed) or letter (A, B, etc.)
            width: Width in characters (approximate)
        """
        if isinstance(col, str):
            col = ord(col.upper()) - ord("A") + 1

        cols = self.sheet.getColumns()
        col_obj = cols.getByIndex(col - 1)
        # Width is in 1/100 mm, convert from character width (approximate: 1 char ~ 2.5 mm)
        col_obj.setPropertyValue("Width", int(width * 250))


def run_fixtures(
    fixtures: Optional[List[str]] = None,
    output_dir: Optional[str] = None,
):
    """
    Run fixture generators.

    Args:
        fixtures: List of fixture filenames to generate. If None, generates all.
        output_dir: Output directory. Defaults to OUTPUT_DIR env var or /output.
    """
    if not HAS_UNO:
        print("ERROR: UNO not available. Cannot generate fixtures.")
        return

    output_dir = output_dir or os.environ.get("OUTPUT_DIR", "/output")
    os.makedirs(output_dir, exist_ok=True)

    to_run = fixtures or list(_FIXTURES.keys())

    print("=" * 60)
    print("PyUNO Fixture Framework")
    print("=" * 60)
    print(f"Output: {output_dir}")
    print(f"Fixtures to generate: {len(to_run)}")
    print()

    for filename in to_run:
        if filename not in _FIXTURES:
            print(f"  WARNING: Unknown fixture '{filename}'")
            continue

        func = _FIXTURES[filename]
        try:
            with FixtureBuilder(filename, output_dir) as builder:
                func(builder)
        except Exception as e:
            print(f"  ERROR generating {filename}: {e}")
            import traceback

            traceback.print_exc()

    print()
    print("=" * 60)
    print("Done!")
    print("=" * 60)


# ============================================================================
# CLI
# ============================================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate XLSX test fixtures using PyUNO"
    )
    parser.add_argument(
        "fixtures",
        nargs="*",
        help="Specific fixtures to generate. If none, generates all.",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output directory",
        default=os.environ.get("OUTPUT_DIR", "/output"),
    )
    parser.add_argument(
        "-l", "--list", action="store_true", help="List available fixtures"
    )

    args = parser.parse_args()

    if args.list:
        print("Available fixtures:")
        for name, func in _FIXTURES.items():
            doc = func.__doc__ or "No description"
            print(f"  {name}: {doc.strip()}")
        sys.exit(0)

    run_fixtures(args.fixtures or None, args.output)
