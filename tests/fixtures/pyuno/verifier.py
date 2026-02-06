#!/usr/bin/env python3
"""
PyUNO Verifier for duke-sheets write tests.

This script opens an XLSX file written by Rust and verifies its contents
using LibreOffice's UNO API. It's used for E2E testing of the writing functionality.

Exit codes:
    0 = All assertions passed
    1 = One or more assertions failed
    2 = Error (file not found, invalid spec, etc.)

Usage:
    python verifier.py <file.xlsx> <spec.json>

The spec.json file contains assertions to verify:
{
    "cells": {
        "A1": {"value": "Hello", "type": "string"},
        "B1": {"value": 42, "type": "number"},
        "C1": {"formula": "=A1&B1"}
    },
    "styles": {
        "A1": {"bold": true, "fill_color": "FF0000"}
    },
    "merges": ["A1:B2", "C1:D1"],
    "comments": {
        "A1": {"text": "Comment text", "author": "Author Name"}
    },
    "sheets": ["Sheet1", "Sheet2"]
}
"""

import json
import sys
import os
from typing import Any, Dict, List, Optional

# UNO imports
try:
    import uno
    from com.sun.star.beans import PropertyValue

    HAS_UNO = True
except ImportError:
    HAS_UNO = False


class VerificationError(Exception):
    """Raised when a verification fails."""

    pass


class Verifier:
    """Verifies XLSX file contents using PyUNO."""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.doc = None
        self.errors: List[str] = []

    def __enter__(self):
        self._connect()
        self._open_document()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.doc:
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

    def _open_document(self):
        """Open the XLSX file."""
        if not os.path.exists(self.filepath):
            raise FileNotFoundError(f"File not found: {self.filepath}")

        url = f"file://{os.path.abspath(self.filepath)}"
        self.doc = self.desktop.loadComponentFromURL(url, "_blank", 0, ())

        if self.doc is None:
            raise RuntimeError(f"Failed to open document: {self.filepath}")

    def _parse_cell(self, cell_ref: str) -> tuple:
        """Parse cell reference like 'A1' into (col, row) zero-indexed."""
        col = 0
        i = 0

        while i < len(cell_ref) and cell_ref[i].isalpha():
            col = col * 26 + (ord(cell_ref[i].upper()) - ord("A") + 1)
            i += 1
        col -= 1

        row = int(cell_ref[i:]) - 1

        return col, row

    def _get_cell(self, sheet, cell_ref: str):
        """Get a cell from a sheet by reference."""
        col, row = self._parse_cell(cell_ref)
        return sheet.getCellByPosition(col, row)

    def _get_sheet(self, name_or_index):
        """Get a sheet by name or index."""
        sheets = self.doc.getSheets()
        if isinstance(name_or_index, int):
            return sheets.getByIndex(name_or_index)
        return sheets.getByName(name_or_index)

    def _add_error(self, message: str):
        """Record a verification error."""
        self.errors.append(message)
        print(f"  FAIL: {message}", file=sys.stderr)

    # =========================================================================
    # Verification Methods
    # =========================================================================

    def verify_sheets(self, expected_sheets: List[str]):
        """Verify sheet names."""
        sheets = self.doc.getSheets()
        actual_names = [
            sheets.getByIndex(i).getName() for i in range(sheets.getCount())
        ]

        if actual_names != expected_sheets:
            self._add_error(
                f"Sheet names mismatch: expected {expected_sheets}, got {actual_names}"
            )

    def verify_cells(self, sheet_name: str, cells: Dict[str, Dict]):
        """Verify cell values and types."""
        try:
            sheet = self._get_sheet(sheet_name)
        except Exception as e:
            self._add_error(f"Sheet '{sheet_name}' not found: {e}")
            return

        for cell_ref, expected in cells.items():
            cell = self._get_cell(sheet, cell_ref)

            # Check value
            if "value" in expected:
                exp_val = expected["value"]
                exp_type = expected.get("type", "auto")

                if exp_type == "string":
                    actual = cell.getString()
                    if actual != exp_val:
                        self._add_error(
                            f"[{sheet_name}!{cell_ref}] String value: expected '{exp_val}', got '{actual}'"
                        )
                elif exp_type == "number":
                    actual = cell.getValue()
                    if abs(actual - exp_val) > 1e-9:
                        self._add_error(
                            f"[{sheet_name}!{cell_ref}] Number value: expected {exp_val}, got {actual}"
                        )
                elif exp_type == "boolean":
                    actual = cell.getValue()
                    expected_num = 1 if exp_val else 0
                    if actual != expected_num:
                        self._add_error(
                            f"[{sheet_name}!{cell_ref}] Boolean value: expected {exp_val}, got {actual}"
                        )
                else:
                    # Auto-detect
                    cell_type = cell.getType()
                    if cell_type == 1:  # VALUE
                        actual = cell.getValue()
                        if isinstance(exp_val, (int, float)):
                            if abs(actual - exp_val) > 1e-9:
                                self._add_error(
                                    f"[{sheet_name}!{cell_ref}] Value: expected {exp_val}, got {actual}"
                                )
                    elif cell_type == 2:  # TEXT
                        actual = cell.getString()
                        if actual != str(exp_val):
                            self._add_error(
                                f"[{sheet_name}!{cell_ref}] Text: expected '{exp_val}', got '{actual}'"
                            )

            # Check formula
            if "formula" in expected:
                actual = cell.getFormula()
                exp_formula = expected["formula"]
                if actual != exp_formula:
                    self._add_error(
                        f"[{sheet_name}!{cell_ref}] Formula: expected '{exp_formula}', got '{actual}'"
                    )

    def verify_styles(self, sheet_name: str, styles: Dict[str, Dict]):
        """Verify cell styles."""
        try:
            sheet = self._get_sheet(sheet_name)
        except Exception as e:
            self._add_error(f"Sheet '{sheet_name}' not found: {e}")
            return

        for cell_ref, expected in styles.items():
            cell = self._get_cell(sheet, cell_ref)

            # Check bold
            if "bold" in expected:
                actual = cell.getPropertyValue("CharWeight") >= 150
                if actual != expected["bold"]:
                    self._add_error(
                        f"[{sheet_name}!{cell_ref}] Bold: expected {expected['bold']}, got {actual}"
                    )

            # Check italic
            if "italic" in expected:
                actual = cell.getPropertyValue("CharPosture") == 2
                if actual != expected["italic"]:
                    self._add_error(
                        f"[{sheet_name}!{cell_ref}] Italic: expected {expected['italic']}, got {actual}"
                    )

            # Check fill color
            if "fill_color" in expected:
                actual = cell.getPropertyValue("CellBackColor")
                exp_color = expected["fill_color"]
                if isinstance(exp_color, str):
                    exp_color = int(exp_color, 16)
                if actual != exp_color:
                    self._add_error(
                        f"[{sheet_name}!{cell_ref}] Fill color: expected {hex(exp_color)}, got {hex(actual)}"
                    )

            # Check font color
            if "font_color" in expected:
                actual = cell.getPropertyValue("CharColor")
                exp_color = expected["font_color"]
                if isinstance(exp_color, str):
                    exp_color = int(exp_color, 16)
                if actual != exp_color:
                    self._add_error(
                        f"[{sheet_name}!{cell_ref}] Font color: expected {hex(exp_color)}, got {hex(actual)}"
                    )

            # Check font size
            if "font_size" in expected:
                actual = cell.getPropertyValue("CharHeight")
                if abs(actual - expected["font_size"]) > 0.1:
                    self._add_error(
                        f"[{sheet_name}!{cell_ref}] Font size: expected {expected['font_size']}, got {actual}"
                    )

    def verify_merges(self, sheet_name: str, expected_merges: List[str]):
        """Verify merged cell regions."""
        try:
            sheet = self._get_sheet(sheet_name)
        except Exception as e:
            self._add_error(f"Sheet '{sheet_name}' not found: {e}")
            return

        # Get actual merges (this is tricky in UNO - need to check each expected range)
        for merge_ref in expected_merges:
            cell_range = sheet.getCellRangeByName(merge_ref)
            # Check if the range is merged
            # In UNO, we check the IsMerged property on the range
            try:
                is_merged = cell_range.getPropertyValue("IsMerged")
                if not is_merged:
                    # Also check via merge cells
                    self._add_error(f"[{sheet_name}] Expected merge at {merge_ref}")
            except Exception:
                self._add_error(f"[{sheet_name}] Could not verify merge at {merge_ref}")

    def verify_comments(self, sheet_name: str, comments: Dict[str, Dict]):
        """Verify cell comments."""
        try:
            sheet = self._get_sheet(sheet_name)
        except Exception as e:
            self._add_error(f"Sheet '{sheet_name}' not found: {e}")
            return

        annotations = sheet.getAnnotations()

        for cell_ref, expected in comments.items():
            col, row = self._parse_cell(cell_ref)

            # Find annotation at this position
            found = False
            for i in range(annotations.getCount()):
                ann = annotations.getByIndex(i)
                pos = ann.getPosition()
                if pos.Column == col and pos.Row == row:
                    found = True

                    # Check text
                    if "text" in expected:
                        actual = ann.getString()
                        if actual != expected["text"]:
                            self._add_error(
                                f"[{sheet_name}!{cell_ref}] Comment text: expected '{expected['text']}', got '{actual}'"
                            )

                    # Check author
                    if "author" in expected:
                        try:
                            actual = ann.getAuthor()
                            if actual != expected["author"]:
                                self._add_error(
                                    f"[{sheet_name}!{cell_ref}] Comment author: expected '{expected['author']}', got '{actual}'"
                                )
                        except Exception:
                            pass  # Author may not be supported

                    break

            if not found:
                self._add_error(f"[{sheet_name}!{cell_ref}] Comment not found")

    def verify_spec(self, spec: Dict):
        """Run all verifications from a spec dict."""
        print(f"Verifying: {self.filepath}")
        print("-" * 60)

        # Verify sheets
        if "sheets" in spec:
            print("Checking sheets...")
            self.verify_sheets(spec["sheets"])

        # Get default sheet name
        default_sheet = spec.get("default_sheet", "Sheet1")
        if "sheets" in spec and spec["sheets"]:
            default_sheet = spec["sheets"][0]

        # Verify cells
        if "cells" in spec:
            print("Checking cells...")
            sheet_cells = spec["cells"]
            if isinstance(sheet_cells, dict):
                # Check if keys are sheet names or cell refs
                first_key = next(iter(sheet_cells.keys()), "")
                if "!" in first_key or first_key.isalpha() or first_key[0].isalpha():
                    # Cell refs in default sheet
                    self.verify_cells(default_sheet, sheet_cells)
                else:
                    # Nested by sheet
                    for sheet_name, cells in sheet_cells.items():
                        self.verify_cells(sheet_name, cells)

        # Verify styles
        if "styles" in spec:
            print("Checking styles...")
            sheet_styles = spec["styles"]
            if isinstance(sheet_styles, dict):
                first_key = next(iter(sheet_styles.keys()), "")
                if first_key and (first_key[0].isalpha() and len(first_key) <= 3):
                    self.verify_styles(default_sheet, sheet_styles)
                else:
                    for sheet_name, styles in sheet_styles.items():
                        self.verify_styles(sheet_name, styles)

        # Verify merges
        if "merges" in spec:
            print("Checking merges...")
            if isinstance(spec["merges"], list):
                self.verify_merges(default_sheet, spec["merges"])
            else:
                for sheet_name, merges in spec["merges"].items():
                    self.verify_merges(sheet_name, merges)

        # Verify comments
        if "comments" in spec:
            print("Checking comments...")
            sheet_comments = spec["comments"]
            if isinstance(sheet_comments, dict):
                first_key = next(iter(sheet_comments.keys()), "")
                if first_key and first_key[0].isalpha():
                    self.verify_comments(default_sheet, sheet_comments)
                else:
                    for sheet_name, comments in sheet_comments.items():
                        self.verify_comments(sheet_name, comments)

        print("-" * 60)
        if self.errors:
            print(f"FAILED: {len(self.errors)} error(s)")
            return False
        else:
            print("PASSED: All assertions passed")
            return True


def main():
    if len(sys.argv) < 2:
        print("Usage: verifier.py <file.xlsx> [spec.json]", file=sys.stderr)
        print("       verifier.py <file.xlsx> --check-opens", file=sys.stderr)
        sys.exit(2)

    filepath = sys.argv[1]

    if not os.path.exists(filepath):
        print(f"Error: File not found: {filepath}", file=sys.stderr)
        sys.exit(2)

    # Simple open check
    if len(sys.argv) == 2 or sys.argv[2] == "--check-opens":
        try:
            with Verifier(filepath) as v:
                print(f"OK: File opens successfully: {filepath}")
                sys.exit(0)
        except Exception as e:
            print(f"Error opening file: {e}", file=sys.stderr)
            sys.exit(1)

    # Full verification with spec
    spec_path = sys.argv[2]
    if not os.path.exists(spec_path):
        print(f"Error: Spec file not found: {spec_path}", file=sys.stderr)
        sys.exit(2)

    with open(spec_path) as f:
        spec = json.load(f)

    try:
        with Verifier(filepath) as v:
            success = v.verify_spec(spec)
            sys.exit(0 if success else 1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback

        traceback.print_exc()
        sys.exit(2)


if __name__ == "__main__":
    main()
