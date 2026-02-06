"""
Border styles fixture - tests cell border formatting.

Tests:
- Border styles (thin, medium, thick, dashed, dotted, double)
- Border colors
- Individual side borders (left, right, top, bottom)
- Box borders (all sides)
- Diagonal borders (where supported)
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder, StyleSpec


@fixture("border_styles.xlsx")
def border_styles(b: FixtureBuilder):
    """Test border styling options."""
    b.set_sheet_name("BorderStyles")

    # === Border styles ===
    b.add_data("A1", "BORDER STYLES")

    styles = [
        ("thin", "Thin border"),
        ("medium", "Medium border"),
        ("thick", "Thick border"),
        ("dashed", "Dashed border"),
        ("dotted", "Dotted border"),
        ("double", "Double border"),
    ]

    for i, (style, desc) in enumerate(styles):
        row = 2 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", f"Style: {style}")
        b.set_cell_style(
            f"B{row}", StyleSpec(border_style=style, border_color=0x000000)
        )

    # === Border colors ===
    b.add_data("A10", "BORDER COLORS")

    colors = [
        ("Black", 0x000000),
        ("Red", 0xFF0000),
        ("Green", 0x00FF00),
        ("Blue", 0x0000FF),
        ("Gray", 0x808080),
    ]

    for i, (name, color) in enumerate(colors):
        row = 11 + i
        b.add_data(f"A{row}", f"{name} border")
        b.add_data(f"B{row}", f"Color: {name}")
        b.set_cell_style(
            f"B{row}", StyleSpec(border_style="medium", border_color=color)
        )

    # === Individual side borders ===
    b.add_data("A18", "INDIVIDUAL BORDERS")

    b.add_data("A19", "Left border only")
    b.add_data("B19", "Left")
    b.set_cell_style(f"B19", StyleSpec(left_border=("medium", 0x000000)))

    b.add_data("A20", "Right border only")
    b.add_data("B20", "Right")
    b.set_cell_style(f"B20", StyleSpec(right_border=("medium", 0x000000)))

    b.add_data("A21", "Top border only")
    b.add_data("B21", "Top")
    b.set_cell_style(f"B21", StyleSpec(top_border=("medium", 0x000000)))

    b.add_data("A22", "Bottom border only")
    b.add_data("B22", "Bottom")
    b.set_cell_style(f"B22", StyleSpec(bottom_border=("medium", 0x000000)))

    # === Mixed borders ===
    b.add_data("A24", "MIXED BORDERS")

    b.add_data("A25", "Top and bottom")
    b.add_data("B25", "TB")
    b.set_cell_style(
        f"B25",
        StyleSpec(
            top_border=("thin", 0x000000),
            bottom_border=("thick", 0x000000),
        ),
    )

    b.add_data("A26", "Left and right")
    b.add_data("B26", "LR")
    b.set_cell_style(
        f"B26",
        StyleSpec(
            left_border=("thin", 0xFF0000),
            right_border=("thin", 0x0000FF),
        ),
    )

    b.add_data("A27", "Different sides")
    b.add_data("B27", "All")
    b.set_cell_style(
        f"B27",
        StyleSpec(
            left_border=("thin", 0xFF0000),
            right_border=("medium", 0x00FF00),
            top_border=("thick", 0x0000FF),
            bottom_border=("dashed", 0x000000),
        ),
    )

    # === Range with box border ===
    b.add_data("A29", "BOX BORDER (RANGE)")

    # Create a range of cells with data
    for row in range(30, 33):
        for col_idx, col in enumerate(["B", "C", "D"]):
            b.add_data(f"{col}{row}", f"{col}{row}")

    # Apply box border to the range
    b.set_range_style(
        "B30:D32", StyleSpec(border_style="medium", border_color=0x000000)
    )


@fixture("border_styles_edge_cases.xlsx")
def border_styles_edge_cases(b: FixtureBuilder):
    """Edge cases for border styling."""
    b.set_sheet_name("EdgeCases")

    # === Very thin vs very thick ===
    b.add_data("A1", "THICKNESS COMPARISON")

    b.add_data("A2", "Thin")
    b.add_data("B2", "Cell")
    b.set_cell_style(f"B2", StyleSpec(border_style="thin", border_color=0x000000))

    b.add_data("A3", "Thick")
    b.add_data("B3", "Cell")
    b.set_cell_style(f"B3", StyleSpec(border_style="thick", border_color=0x000000))

    # === Border with fill ===
    b.add_data("A5", "BORDER + FILL")

    b.add_data("A6", "Red fill, black border")
    b.add_data("B6", "Combined")
    b.set_cell_style(
        f"B6",
        StyleSpec(
            fill_color=0xFF0000,
            border_style="medium",
            border_color=0x000000,
        ),
    )

    b.add_data("A7", "Yellow fill, blue border")
    b.add_data("B7", "Combined")
    b.set_cell_style(
        f"B7",
        StyleSpec(
            fill_color=0xFFFF00,
            border_style="medium",
            border_color=0x0000FF,
        ),
    )

    # === Adjacent cells with borders ===
    b.add_data("A9", "ADJACENT BORDERS")

    b.add_data("B10", "Cell 1")
    b.add_data("C10", "Cell 2")
    b.set_cell_style(f"B10", StyleSpec(border_style="thin", border_color=0x000000))
    b.set_cell_style(f"C10", StyleSpec(border_style="thin", border_color=0xFF0000))
