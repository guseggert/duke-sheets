"""
Fill styles fixture - tests background/fill formatting.

Tests:
- Solid color fills (various colors)
- Pattern fills (gray125, etc.) - Note: LibreOffice support may be limited
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder, StyleSpec


@fixture("fill_styles.xlsx")
def fill_styles(b: FixtureBuilder):
    """Test fill/background styling options."""
    b.set_sheet_name("FillStyles")

    # === Solid fills ===
    b.add_data("A1", "SOLID FILLS")

    colors = [
        ("Red", 0xFF0000),
        ("Green", 0x00FF00),
        ("Blue", 0x0000FF),
        ("Yellow", 0xFFFF00),
        ("Cyan", 0x00FFFF),
        ("Magenta", 0xFF00FF),
        ("Orange", 0xFFA500),
        ("Purple", 0x800080),
        ("Pink", 0xFFC0CB),
        ("Light Gray", 0xD3D3D3),
        ("Dark Gray", 0x404040),
        ("White", 0xFFFFFF),
        ("Black", 0x000000),
    ]

    for i, (name, color) in enumerate(colors):
        row = 2 + i
        b.add_data(f"A{row}", name)
        b.add_data(f"B{row}", f"Fill: {name}")
        b.set_cell_style(f"B{row}", StyleSpec(fill_color=color))
        # For black fill, use white text
        if color == 0x000000:
            b.set_cell_style(
                f"B{row}", StyleSpec(fill_color=color, font_color=0xFFFFFF)
            )

    # === Pastel colors ===
    b.add_data("A16", "PASTEL FILLS")

    pastels = [
        ("Light Red", 0xFFCCCC),
        ("Light Green", 0xCCFFCC),
        ("Light Blue", 0xCCCCFF),
        ("Light Yellow", 0xFFFFCC),
        ("Light Cyan", 0xCCFFFF),
        ("Light Magenta", 0xFFCCFF),
    ]

    for i, (name, color) in enumerate(pastels):
        row = 17 + i
        b.add_data(f"A{row}", name)
        b.add_data(f"B{row}", f"Fill: {name}")
        b.set_cell_style(f"B{row}", StyleSpec(fill_color=color))

    # === Fill + Font color combinations ===
    b.add_data("A24", "FILL + FONT COMBINATIONS")

    combos = [
        ("White on Red", 0xFF0000, 0xFFFFFF),
        ("White on Blue", 0x0000FF, 0xFFFFFF),
        ("Black on Yellow", 0xFFFF00, 0x000000),
        ("Yellow on Black", 0x000000, 0xFFFF00),
        ("Red on White", 0xFFFFFF, 0xFF0000),
    ]

    for i, (name, fill, font) in enumerate(combos):
        row = 25 + i
        b.add_data(f"A{row}", name)
        b.add_data(f"B{row}", f"Text with {name}")
        b.set_cell_style(f"B{row}", StyleSpec(fill_color=fill, font_color=font))
