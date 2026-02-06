"""
Comments fixture - tests cell comments/annotations.

Tests:
- Basic comments
- Comments with author
- Long comments
- Comments on styled cells
- Comments on merged cells
- Unicode in comments
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder, StyleSpec


@fixture("comments.xlsx")
def comments(b: FixtureBuilder):
    """Test cell comment options."""
    b.set_sheet_name("Comments")

    b.set_column_width("A", 25)
    b.set_column_width("B", 20)

    # === Basic comments ===
    b.add_data("A1", "BASIC COMMENTS")

    b.add_data("A2", "Simple comment")
    b.add_data("B2", "Has comment")
    b.add_comment("B2", "This is a simple comment.")

    b.add_data("A3", "With author")
    b.add_data("B3", "Has comment")
    b.add_comment("B3", "This comment has an author.", author="Test Author")

    b.add_data("A4", "Multi-line comment")
    b.add_data("B4", "Has comment")
    b.add_comment(
        "B4",
        "Line 1 of the comment.\nLine 2 of the comment.\nLine 3 of the comment.",
    )

    # === Long comments ===
    b.add_data("A7", "LONG COMMENTS")

    b.add_data("A8", "Long text")
    b.add_data("B8", "Has comment")
    long_text = (
        "This is a very long comment that contains a lot of text. "
        "It should test how the system handles comments that exceed typical lengths. "
        "Comments in Excel can contain quite a bit of text, and it's important to "
        "ensure that all of the content is properly preserved when reading and writing. "
        "This comment continues with more text to make it even longer."
    )
    b.add_comment("B8", long_text)

    # === Comments on styled cells ===
    b.add_data("A11", "COMMENTS + STYLING")

    b.add_data("A12", "Bold cell with comment")
    b.add_data("B12", "Bold + Comment")
    b.set_cell_style("B12", StyleSpec(bold=True))
    b.add_comment("B12", "This cell is bold and has a comment.")

    b.add_data("A13", "Colored cell with comment")
    b.add_data("B13", "Colored + Comment")
    b.set_cell_style("B13", StyleSpec(fill_color=0xFFFF00))
    b.add_comment("B13", "This cell has a yellow background and a comment.")

    b.add_data("A14", "Full styling + comment")
    b.add_data("B14", "Styled + Comment")
    b.set_cell_style(
        "B14",
        StyleSpec(
            bold=True,
            italic=True,
            font_color=0x0000FF,
            fill_color=0xFFCCCC,
            border_style="thin",
            border_color=0x000000,
        ),
    )
    b.add_comment("B14", "This cell has full styling and a comment.", author="Stylist")

    # === Multiple comments in range ===
    b.add_data("A17", "MULTIPLE COMMENTS")

    comments_data = [
        ("B18", "Cell 1", "Comment for cell 1"),
        ("B19", "Cell 2", "Comment for cell 2"),
        ("B20", "Cell 3", "Comment for cell 3"),
        ("B21", "Cell 4", "Comment for cell 4"),
    ]
    for cell, value, comment in comments_data:
        b.add_data(cell, value)
        b.add_comment(cell, comment)


@fixture("comments_unicode.xlsx")
def comments_unicode(b: FixtureBuilder):
    """Test unicode characters in comments."""
    b.set_sheet_name("Unicode")

    b.set_column_width("A", 30)
    b.set_column_width("B", 15)

    # === Unicode in comments ===
    b.add_data("A1", "UNICODE IN COMMENTS")

    unicode_tests = [
        ("A2", "B2", "German", "Umlauts: aou"),
        ("A3", "B3", "French", "Accents: cafe, ecole"),
        ("A4", "B4", "Spanish", "Tildes: manana, nino"),
        ("A5", "B5", "Greek", "Greek: alpha, beta, gamma"),
        ("A6", "B6", "Cyrillic", "Russian: Privet, Mir"),
        ("A7", "B7", "Japanese", "Hiragana: Arigatou"),
        ("A8", "B8", "Chinese", "Hanzi: Hello"),
        ("A9", "B9", "Emoji", "Emojis: Smile, Heart, Star"),
        ("A10", "B10", "Math", "Math symbols: Sum, Integral, Infinity"),
    ]

    for label_cell, value_cell, label, comment in unicode_tests:
        b.add_data(label_cell, label)
        b.add_data(value_cell, label)
        b.add_comment(value_cell, comment)

    # === Special characters ===
    b.add_data("A13", "SPECIAL CHARACTERS")

    special_tests = [
        ("A14", "B14", "Quotes", "Single ' and double \" quotes"),
        ("A15", "B15", "Angles", "Less < and greater > than"),
        ("A16", "B16", "Ampersand", "Ampersand & character"),
        ("A17", "B17", "Newlines", "Line1\nLine2\nLine3"),
        ("A18", "B18", "Tabs", "Col1\tCol2\tCol3"),
    ]

    for label_cell, value_cell, label, comment in special_tests:
        b.add_data(label_cell, label)
        b.add_data(value_cell, label)
        b.add_comment(value_cell, comment)


@fixture("comments_edge_cases.xlsx")
def comments_edge_cases(b: FixtureBuilder):
    """Edge cases for comments."""
    b.set_sheet_name("EdgeCases")

    b.set_column_width("A", 30)
    b.set_column_width("B", 15)

    # === Empty-ish comments ===
    b.add_data("A1", "EDGE CASES")

    b.add_data("A2", "Single character")
    b.add_data("B2", "Value")
    b.add_comment("B2", "X")

    b.add_data("A3", "Only whitespace")
    b.add_data("B3", "Value")
    b.add_comment("B3", "   ")

    b.add_data("A4", "Only newlines")
    b.add_data("B4", "Value")
    b.add_comment("B4", "\n\n\n")

    # === Comment on different cell types ===
    b.add_data("A7", "DIFFERENT CELL TYPES")

    b.add_data("A8", "Number cell")
    b.add_data("B8", 12345)
    b.add_comment("B8", "Comment on a number")

    b.add_data("A9", "Formula cell")
    b.add_formula("B9", "=1+1")
    b.add_comment("B9", "Comment on a formula")

    b.add_data("A10", "Date cell")
    b.add_data("B10", 45366)  # Excel date serial
    b.set_cell_style("B10", StyleSpec(number_format="YYYY-MM-DD"))
    b.add_comment("B10", "Comment on a date")

    b.add_data("A11", "Boolean cell")
    b.add_data("B11", True)
    b.add_comment("B11", "Comment on a boolean")

    # === Adjacent cells with comments ===
    b.add_data("A14", "ADJACENT COMMENTS")

    for col in ["B", "C", "D", "E"]:
        b.add_data(f"{col}15", col)
        b.add_comment(f"{col}15", f"Comment for column {col}")

    # === Comment with author containing special characters ===
    b.add_data("A18", "AUTHOR SPECIAL CHARS")
    b.add_data("B18", "Value")
    b.add_comment("B18", "Comment text", author="John O'Brien")
