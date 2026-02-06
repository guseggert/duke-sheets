"""
Font styles fixture - tests all font formatting options.

Tests:
- Bold, italic, underline (5 types), strikethrough
- Font colors (RGB, various colors)
- Font sizes (8pt to 24pt)
- Font names (Arial, Times New Roman, Courier New)
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder, StyleSpec


@fixture("font_styles.xlsx")
def font_styles(b: FixtureBuilder):
    """Test all font styling options."""
    b.set_sheet_name("FontStyles")

    # === Basic formatting ===
    b.add_data("A1", "BASIC FORMATTING")

    b.add_data("A2", "Bold")
    b.add_data("B2", "Bold Text")
    b.set_cell_style("B2", StyleSpec(bold=True))

    b.add_data("A3", "Italic")
    b.add_data("B3", "Italic Text")
    b.set_cell_style("B3", StyleSpec(italic=True))

    b.add_data("A4", "Bold + Italic")
    b.add_data("B4", "Bold Italic Text")
    b.set_cell_style("B4", StyleSpec(bold=True, italic=True))

    b.add_data("A5", "Strikethrough")
    b.add_data("B5", "Strikethrough Text")
    b.set_cell_style("B5", StyleSpec(strikethrough=True))

    # === Underline types ===
    b.add_data("A7", "UNDERLINE TYPES")

    b.add_data("A8", "Single")
    b.add_data("B8", "Single Underline")
    b.set_cell_style("B8", StyleSpec(underline="single"))

    b.add_data("A9", "Double")
    b.add_data("B9", "Double Underline")
    b.set_cell_style("B9", StyleSpec(underline="double"))

    b.add_data("A10", "Single Accounting")
    b.add_data("B10", "Single Accounting")
    b.set_cell_style("B10", StyleSpec(underline="singleAccounting"))

    b.add_data("A11", "Double Accounting")
    b.add_data("B11", "Double Accounting")
    b.set_cell_style("B11", StyleSpec(underline="doubleAccounting"))

    # === Font colors ===
    b.add_data("A13", "FONT COLORS")

    colors = [
        ("Red", 0xFF0000),
        ("Green", 0x00FF00),
        ("Blue", 0x0000FF),
        ("Yellow", 0xFFFF00),
        ("Purple", 0x800080),
        ("Orange", 0xFFA500),
        ("Black", 0x000000),
        ("Gray", 0x808080),
    ]

    for i, (name, color) in enumerate(colors):
        row = 14 + i
        b.add_data(f"A{row}", name)
        b.add_data(f"B{row}", f"{name} Text")
        b.set_cell_style(f"B{row}", StyleSpec(font_color=color))

    # === Font sizes ===
    b.add_data("A23", "FONT SIZES")

    sizes = [8, 10, 11, 12, 14, 16, 18, 20, 24]
    for i, size in enumerate(sizes):
        row = 24 + i
        b.add_data(f"A{row}", f"{size}pt")
        b.add_data(f"B{row}", f"Size {size}")
        b.set_cell_style(f"B{row}", StyleSpec(font_size=size))

    # === Font names ===
    b.add_data("A34", "FONT NAMES")

    fonts = ["Arial", "Times New Roman", "Courier New", "Verdana", "Georgia"]
    for i, font in enumerate(fonts):
        row = 35 + i
        b.add_data(f"A{row}", font)
        b.add_data(f"B{row}", f"Font: {font}")
        b.set_cell_style(f"B{row}", StyleSpec(font_name=font))

    # === Combinations ===
    b.add_data("A41", "COMBINATIONS")

    b.add_data("A42", "All effects")
    b.add_data("B42", "Bold + Italic + Underline + Color")
    b.set_cell_style(
        "B42",
        StyleSpec(
            bold=True,
            italic=True,
            underline="single",
            font_color=0x0000FF,
            font_size=14,
        ),
    )
