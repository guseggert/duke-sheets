"""
Alignment fixture - tests text alignment and orientation.

Tests:
- Horizontal alignment (left, center, right, justify, fill)
- Vertical alignment (top, center, bottom)
- Text wrapping
- Shrink to fit
- Text rotation (-90 to 90 degrees, plus vertical stacking)
- Indent levels
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder, StyleSpec


@fixture("alignment.xlsx")
def alignment(b: FixtureBuilder):
    """Test text alignment options."""
    b.set_sheet_name("Alignment")

    # Set column widths for visibility
    b.set_column_width("B", 20)
    b.set_column_width("C", 15)

    # === Horizontal alignment ===
    b.add_data("A1", "HORIZONTAL ALIGNMENT")

    alignments = [
        ("left", "Left aligned text"),
        ("center", "Centered text"),
        ("right", "Right aligned text"),
        ("justify", "Justified text that spans the full width of the cell"),
        ("fill", "Fill"),
        ("general", "General (default)"),
        ("distributed", "Distributed text"),
    ]

    for i, (align, desc) in enumerate(alignments):
        row = 2 + i
        b.add_data(f"A{row}", align)
        b.add_data(f"B{row}", desc)
        b.set_cell_style(f"B{row}", StyleSpec(horizontal=align))

    # === Vertical alignment ===
    b.add_data("A11", "VERTICAL ALIGNMENT")

    # Set row heights for visibility
    for row in range(12, 17):
        b.set_row_height(row, 40)

    v_alignments = [
        ("top", "Top aligned"),
        ("center", "Center aligned"),
        ("bottom", "Bottom aligned"),
        ("justify", "Justified vertically"),
    ]

    for i, (align, desc) in enumerate(v_alignments):
        row = 12 + i
        b.add_data(f"A{row}", align)
        b.add_data(f"B{row}", desc)
        b.set_cell_style(f"B{row}", StyleSpec(vertical=align))

    # === Combined horizontal and vertical ===
    b.add_data("A18", "COMBINED H+V ALIGNMENT")
    b.set_row_height(19, 50)
    b.add_data("A19", "Center/Center")
    b.add_data("B19", "Centered both ways")
    b.set_cell_style(f"B19", StyleSpec(horizontal="center", vertical="center"))

    b.set_row_height(20, 50)
    b.add_data("A20", "Right/Bottom")
    b.add_data("B20", "Right and bottom")
    b.set_cell_style(f"B20", StyleSpec(horizontal="right", vertical="bottom"))

    # === Text wrapping ===
    b.add_data("A22", "TEXT WRAPPING")
    b.set_row_height(23, 60)
    b.add_data("A23", "Wrap text")
    b.add_data(
        "B23",
        "This is a long text that should wrap to multiple lines within the cell boundaries",
    )
    b.set_cell_style(f"B23", StyleSpec(wrap_text=True))

    b.set_row_height(24, 30)
    b.add_data("A24", "No wrap")
    b.add_data(
        "B24",
        "This is a long text without wrapping that will overflow",
    )

    # === Shrink to fit ===
    b.add_data("A26", "SHRINK TO FIT")
    b.add_data("A27", "Shrink to fit")
    b.add_data("B27", "This long text should shrink to fit in the cell")
    b.set_cell_style(f"B27", StyleSpec(shrink_to_fit=True))

    # === Indent ===
    b.add_data("A29", "INDENT LEVELS")

    for i in range(5):
        row = 30 + i
        b.add_data(f"A{row}", f"Indent {i}")
        b.add_data(f"B{row}", f"Text with indent level {i}")
        b.set_cell_style(f"B{row}", StyleSpec(indent=i))


@fixture("alignment_rotation.xlsx")
def alignment_rotation(b: FixtureBuilder):
    """Test text rotation options."""
    b.set_sheet_name("Rotation")

    # Set row heights for rotated text visibility
    for row in range(1, 20):
        b.set_row_height(row, 60)

    b.set_column_width("B", 15)

    # === Positive rotation (counterclockwise) ===
    b.add_data("A1", "POSITIVE ROTATION")

    rotations = [0, 15, 30, 45, 60, 75, 90]
    for i, angle in enumerate(rotations):
        row = 2 + i
        b.add_data(f"A{row}", f"{angle} degrees")
        b.add_data(f"B{row}", f"Text at {angle}")
        b.set_cell_style(f"B{row}", StyleSpec(rotation=angle))

    # === Negative rotation (clockwise) ===
    b.add_data("A10", "NEGATIVE ROTATION")

    neg_rotations = [-15, -30, -45, -60, -75, -90]
    for i, angle in enumerate(neg_rotations):
        row = 11 + i
        b.add_data(f"A{row}", f"{angle} degrees")
        b.add_data(f"B{row}", f"Text at {angle}")
        b.set_cell_style(f"B{row}", StyleSpec(rotation=angle))

    # === Vertical stacking (special value 255) ===
    b.add_data("A18", "VERTICAL STACKING")
    b.add_data("A19", "Vertical (255)")
    b.add_data("B19", "Stacked")
    b.set_cell_style(f"B19", StyleSpec(rotation=255))


@fixture("alignment_edge_cases.xlsx")
def alignment_edge_cases(b: FixtureBuilder):
    """Edge cases for alignment."""
    b.set_sheet_name("EdgeCases")

    # === Wrap + alignment ===
    b.add_data("A1", "WRAP + ALIGNMENT")
    b.set_row_height(2, 80)
    b.set_column_width("B", 20)

    b.add_data("A2", "Wrap + center")
    b.add_data(
        "B2", "Long wrapped text that should be centered horizontally and vertically"
    )
    b.set_cell_style(
        f"B2",
        StyleSpec(wrap_text=True, horizontal="center", vertical="center"),
    )

    # === Rotation + alignment ===
    b.set_row_height(4, 80)
    b.add_data("A4", "Rotation + alignment")
    b.add_data("B4", "Rotated centered text")
    b.set_cell_style(
        f"B4",
        StyleSpec(rotation=45, horizontal="center", vertical="center"),
    )

    # === Indent + alignment ===
    b.add_data("A6", "Indent + right align")
    b.add_data("B6", "Right aligned with indent")
    # Note: indent with right alignment may behave differently
    b.set_cell_style(f"B6", StyleSpec(indent=2, horizontal="right"))

    # === All properties combined ===
    b.set_row_height(8, 100)
    b.add_data("A8", "Combined properties")
    b.add_data("B8", "All alignment properties combined")
    b.set_cell_style(
        f"B8",
        StyleSpec(
            horizontal="center",
            vertical="center",
            wrap_text=True,
            indent=1,
        ),
    )
