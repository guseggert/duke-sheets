#!/usr/bin/env python3
"""
PyUNO Fixture Framework for generating XLSX test files.

This provides a clean abstraction for creating test fixtures using LibreOffice's
UNO API. It handles connection management, document creation, and common patterns.

Usage:
    # Define a fixture
    @fixture("my_test.xlsx")
    def create_my_test(builder):
        builder.add_data("A1:A5", [10, 20, 30, 40, 50])
        builder.add_conditional_format(
            range="A1:A5",
            condition=("greater_than", 25),
            style={
                "bold": True,
                "fill_color": 0xFFFF00,
            }
        )

    # Run all fixtures
    if __name__ == "__main__":
        run_fixtures()
"""

import os
import sys
from typing import Dict, List, Any, Callable, Optional, Tuple, Union
from dataclasses import dataclass, field
from contextlib import contextmanager

# UNO imports
try:
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.sheet.ConditionOperator import GREATER, LESS, EQUAL, BETWEEN, NONE

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


@dataclass
class StyleSpec:
    """Specification for a cell style."""

    # Font
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    font_color: Optional[int] = None  # RGB as int, e.g., 0xFF0000
    font_size: Optional[float] = None
    font_name: Optional[str] = None

    # Fill
    fill_color: Optional[int] = None  # RGB as int

    # Alignment
    horizontal: Optional[str] = None  # "left", "center", "right"
    vertical: Optional[str] = None  # "top", "center", "bottom"
    wrap_text: bool = False
    rotation: int = 0
    indent: int = 0

    # Border
    border_style: Optional[str] = None  # "thin", "medium", "thick"
    border_color: Optional[int] = None

    # Number format
    number_format: Optional[str] = None  # e.g., "0.00%", "#,##0"

    # Protection
    locked: Optional[bool] = None
    hidden: Optional[bool] = None


class FixtureBuilder:
    """
    Builder for creating XLSX test fixtures with LibreOffice.

    Example:
        with FixtureBuilder("test.xlsx") as builder:
            builder.add_data("A1:A5", [1, 2, 3, 4, 5])
            builder.add_conditional_format(
                range="A1:A5",
                condition=("greater_than", 3),
                style=StyleSpec(bold=True, fill_color=0xFFFF00)
            )
    """

    # Condition type mapping
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
    }

    # Vertical alignment mapping
    VERTICAL_ALIGN = {
        "standard": 0,
        "top": 1,
        "center": 2,
        "bottom": 3,
    }

    # Border style mapping
    BORDER_STYLES = {
        "none": 0,
        "thin": 1,
        "medium": 2,
        "thick": 3,
    }

    def __init__(self, filename: str, output_dir: Optional[str] = None):
        self.filename = filename
        self.output_dir = output_dir or os.environ.get("OUTPUT_DIR", "/output")
        self.doc = None
        self.sheet = None
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

    def _make_prop(self, name: str, value: Any) -> PropertyValue:
        """Create a UNO PropertyValue."""
        pv = PropertyValue()
        pv.Name = name
        pv.Value = value
        return pv

    def set_sheet_name(self, name: str):
        """Set the active sheet's name."""
        self.sheet.setName(name)

    def add_data(self, start_cell: str, data: Union[List, List[List]]):
        """
        Add data to the sheet.

        Args:
            start_cell: Starting cell like "A1"
            data: List of values (column) or list of lists (rows)
        """
        # Parse start cell
        col, row = self._parse_cell(start_cell)

        # Normalize to list of lists
        if data and not isinstance(data[0], (list, tuple)):
            data = [[v] for v in data]  # Single column

        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                cell = self.sheet.getCellByPosition(col + col_idx, row + row_idx)
                if isinstance(value, str):
                    cell.setString(value)
                else:
                    cell.setValue(value)

    def _parse_cell(self, cell_ref: str) -> Tuple[int, int]:
        """Parse cell reference like 'A1' into (col, row) zero-indexed."""
        col = 0
        row = 0
        i = 0

        # Parse column letters
        while i < len(cell_ref) and cell_ref[i].isalpha():
            col = col * 26 + (ord(cell_ref[i].upper()) - ord("A") + 1)
            i += 1
        col -= 1  # Zero-indexed

        # Parse row number
        row = int(cell_ref[i:]) - 1  # Zero-indexed

        return col, row

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

        # Create the cell style
        style_name = self._create_style(style)

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

    def _create_style(self, spec: StyleSpec) -> str:
        """Create a cell style from a StyleSpec and return its name."""
        self._style_counter += 1
        style_name = f"CF_Style_{self._style_counter}"

        # Get style families
        style_families = self.doc.getStyleFamilies()
        cell_styles = style_families.getByName("CellStyles")

        # Create style
        style = self.doc.createInstance("com.sun.star.style.CellStyle")
        cell_styles.insertByName(style_name, style)

        # Apply font properties
        if spec.bold:
            style.setPropertyValue("CharWeight", 150)  # BOLD
        if spec.italic:
            style.setPropertyValue("CharPosture", 2)  # ITALIC
        if spec.underline:
            style.setPropertyValue("CharUnderline", 1)  # SINGLE
        if spec.strikethrough:
            style.setPropertyValue("CharStrikeout", 1)
        if spec.font_color is not None:
            style.setPropertyValue("CharColor", spec.font_color)
        if spec.font_size is not None:
            style.setPropertyValue("CharHeight", spec.font_size)
        if spec.font_name is not None:
            style.setPropertyValue("CharFontName", spec.font_name)

        # Apply fill
        if spec.fill_color is not None:
            style.setPropertyValue("CellBackColor", spec.fill_color)

        # Apply alignment
        if spec.horizontal is not None:
            h_val = self.HORIZONTAL_ALIGN.get(spec.horizontal, 0)
            style.setPropertyValue("HoriJustify", h_val)
        if spec.vertical is not None:
            v_val = self.VERTICAL_ALIGN.get(spec.vertical, 0)
            style.setPropertyValue("VertJustify", v_val)
        if spec.wrap_text:
            style.setPropertyValue("IsTextWrapped", True)
        if spec.rotation != 0:
            style.setPropertyValue(
                "RotateAngle", spec.rotation * 100
            )  # In 1/100 degree
        if spec.indent > 0:
            style.setPropertyValue("ParaIndent", spec.indent * 200)  # Approximate

        # Apply borders
        if spec.border_style is not None:
            border_line = uno.createUnoStruct("com.sun.star.table.BorderLine2")
            border_line.LineStyle = self.BORDER_STYLES.get(spec.border_style, 1)
            border_line.LineWidth = 50 if spec.border_style == "thin" else 100
            border_line.Color = spec.border_color or 0x000000

            for side in ["TopBorder", "BottomBorder", "LeftBorder", "RightBorder"]:
                style.setPropertyValue(side, border_line)

        # Apply number format
        if spec.number_format is not None:
            number_formats = self.doc.getNumberFormats()
            locale = uno.createUnoStruct("com.sun.star.lang.Locale")
            fmt_id = number_formats.queryKey(spec.number_format, locale, False)
            if fmt_id == -1:
                fmt_id = number_formats.addNew(spec.number_format, locale)
            style.setPropertyValue("NumberFormat", fmt_id)

        return style_name


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
# Built-in fixtures using the framework
# ============================================================================


@fixture("pyuno_dxf_font.xlsx")
def fixture_font(b: FixtureBuilder):
    """DXF with font formatting."""
    b.set_sheet_name("FontTest")
    b.add_data("A1", [(i + 1) * 10 for i in range(10)])
    b.add_conditional_format(
        range="A1:A10",
        condition=("greater_than", 50),
        style={"bold": True, "font_color": 0xFF0000},
    )


@fixture("pyuno_dxf_fill.xlsx")
def fixture_fill(b: FixtureBuilder):
    """DXF with fill color."""
    b.set_sheet_name("FillTest")
    b.add_data("A1", [(i + 1) * 10 for i in range(10)])
    b.add_conditional_format(
        range="A1:A10", condition=("greater_than", 50), style={"fill_color": 0xFFFF00}
    )


@fixture("pyuno_dxf_border.xlsx")
def fixture_border(b: FixtureBuilder):
    """DXF with borders."""
    b.set_sheet_name("BorderTest")
    b.add_data("A1", [(i + 1) * 10 for i in range(10)])
    b.add_conditional_format(
        range="A1:A10",
        condition=("greater_than", 50),
        style={"border_style": "thin", "border_color": 0x0000FF},
    )


@fixture("pyuno_dxf_alignment.xlsx")
def fixture_alignment(b: FixtureBuilder):
    """DXF with alignment."""
    b.set_sheet_name("AlignTest")
    b.add_data("A1", [f"Text {i + 1}" for i in range(10)])
    b.add_conditional_format(
        range="A1:A10",
        condition=("equal", '"Text 5"'),
        style={"horizontal": "center", "vertical": "center", "wrap_text": True},
    )


@fixture("pyuno_dxf_numfmt.xlsx")
def fixture_numfmt(b: FixtureBuilder):
    """DXF with number format."""
    b.set_sheet_name("NumFmtTest")
    b.add_data("A1", [(i + 1) * 0.123 for i in range(10)])
    b.add_conditional_format(
        range="A1:A10",
        condition=("greater_than", 0.5),
        style={"number_format": "0.00%"},
    )


@fixture("pyuno_dxf_full.xlsx")
def fixture_full(b: FixtureBuilder):
    """DXF with all formatting properties."""
    b.set_sheet_name("FullTest")
    b.add_data("A1", [(i + 1) * 10 for i in range(10)])
    b.add_conditional_format(
        range="A1:A10",
        condition=("greater_than", 50),
        style=StyleSpec(
            bold=True,
            italic=True,
            font_color=0xFF0000,
            fill_color=0xFFFF00,
            horizontal="center",
            vertical="center",
            wrap_text=True,
            border_style="thin",
            border_color=0x0000FF,
        ),
    )


@fixture("pyuno_dxf_multiple_rules.xlsx")
def fixture_multiple_rules(b: FixtureBuilder):
    """Multiple CF rules on same range."""
    b.set_sheet_name("MultipleRules")
    b.add_data("A1", [(i + 1) * 10 for i in range(10)])

    # High values = red
    b.add_conditional_format(
        range="A1:A10",
        condition=("greater_than", 70),
        style={"fill_color": 0xFF0000, "bold": True},
    )
    # Medium values = yellow
    b.add_conditional_format(
        range="A1:A10", condition=("greater_than", 40), style={"fill_color": 0xFFFF00}
    )
    # Low values = green
    b.add_conditional_format(
        range="A1:A10", condition=("greater_than", 20), style={"fill_color": 0x00FF00}
    )


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
