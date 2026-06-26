"""Typing helpers for duke_sheets."""

from typing import List, Literal, TypedDict


class ColorInput(TypedDict, total=False):
    color_type: Literal["auto", "rgb", "argb", "theme", "indexed"]
    hex: str
    r: int
    g: int
    b: int
    a: int
    theme_index: int
    tint: int
    palette_index: int


class FontStyleInput(TypedDict, total=False):
    name: str
    size: float
    bold: bool
    italic: bool
    underline: Literal["none", "single", "double", "singleAccounting", "doubleAccounting"]
    strikethrough: bool
    color: ColorInput
    vertical_align: Literal["baseline", "superscript", "subscript"]
    family: int
    charset: int
    scheme: str


class GradientStopInput(TypedDict):
    position: float
    color: ColorInput


class FillStyleInput(TypedDict, total=False):
    fill_type: Literal["none", "solid", "pattern", "gradient"]
    color: ColorInput
    pattern: str
    foreground: ColorInput
    background: ColorInput
    gradient_type: Literal["linear", "path"]
    angle: float
    stops: List[GradientStopInput]


class BorderEdgeInput(TypedDict, total=False):
    style: Literal[
        "none",
        "thin",
        "medium",
        "thick",
        "dashed",
        "dotted",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
    ]
    color: ColorInput


class BorderStyleInput(TypedDict, total=False):
    left: BorderEdgeInput
    right: BorderEdgeInput
    top: BorderEdgeInput
    bottom: BorderEdgeInput
    diagonal: BorderEdgeInput
    diagonal_direction: Literal["none", "down", "up", "both"]


class AlignmentInput(TypedDict, total=False):
    horizontal: Literal[
        "general",
        "left",
        "center",
        "right",
        "fill",
        "justify",
        "centerContinuous",
        "distributed",
    ]
    vertical: Literal["top", "center", "bottom", "justify", "distributed"]
    wrap_text: bool
    shrink_to_fit: bool
    indent: int
    rotation: int
    reading_order: Literal["contextDependent", "leftToRight", "rightToLeft"]


class NumberFormatInput(TypedDict, total=False):
    format_type: Literal["general", "builtin", "custom"]
    id: int
    format_string: str


class CellProtectionInput(TypedDict, total=False):
    locked: bool
    hidden: bool


class StyleInput(TypedDict, total=False):
    font: FontStyleInput
    fill: FillStyleInput
    border: BorderStyleInput
    alignment: AlignmentInput
    number_format: NumberFormatInput
    protection: CellProtectionInput


class WorkbookProtectionInput(TypedDict, total=False):
    structure: bool
    windows: bool
    password: str
    password_hash: int


class SheetProtectionInput(TypedDict, total=False):
    protected: bool
    password: str
    password_hash: int
    select_locked_cells: bool
    select_unlocked_cells: bool
    format_cells: bool
    format_columns: bool
    format_rows: bool
    insert_columns: bool
    insert_rows: bool
    insert_hyperlinks: bool
    delete_columns: bool
    delete_rows: bool
    sort: bool
    auto_filter: bool
    pivot_tables: bool


class ProtectedRangeInput(TypedDict, total=False):
    name: str
    ranges: List[str]
    password: str
    password_hash: int
    security_descriptor: str


__all__ = [
    "AlignmentInput",
    "BorderEdgeInput",
    "BorderStyleInput",
    "CellProtectionInput",
    "ColorInput",
    "FillStyleInput",
    "FontStyleInput",
    "GradientStopInput",
    "NumberFormatInput",
    "ProtectedRangeInput",
    "SheetProtectionInput",
    "StyleInput",
    "WorkbookProtectionInput",
]
