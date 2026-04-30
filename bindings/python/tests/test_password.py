"""
End-to-end tests for password-protected save/open round-trips through
the Python bindings.
"""

import os
import tempfile

import pytest


PASSWORD = "duke-test-pw"


def _build_sample():
    import duke_sheets

    wb = duke_sheets.Workbook()
    sheet = wb.get_sheet(0)
    sheet.set_cell("A1", "hello")
    sheet.set_cell("B1", 42.0)
    sheet.set_cell("A2", 3.14)
    sheet.set_cell("B2", True)
    return wb


def _cell_value(sheet, addr):
    cell = sheet.get_cell(addr)
    return cell.to_python() if cell is not None else None


def _round_trip(extension, profile=None, **kwargs):
    import duke_sheets

    wb = _build_sample()
    with tempfile.NamedTemporaryFile(suffix=extension, delete=False) as f:
        path = f.name
    try:
        wb.save_with_password(path, PASSWORD, profile, **kwargs)
        opened = duke_sheets.Workbook.open_with_password(path, PASSWORD)
        sheet = opened.get_sheet(0)
        assert _cell_value(sheet, "A1") == "hello"
        b1 = _cell_value(sheet, "B1")
        assert b1 == 42.0 or b1 == 42, f"B1 was {b1!r}"
    finally:
        os.unlink(path)


class TestXlsxPassword:
    def test_default_profile(self):
        _round_trip(".xlsx")

    def test_agile_profile(self):
        _round_trip(".xlsx", profile="agile", key_bits=256)

    def test_standard_profile(self):
        _round_trip(".xlsx", profile="standard")


class TestXlsPassword:
    def test_default_profile(self):
        _round_trip(".xls")

    def test_rc4_cryptoapi_128(self):
        _round_trip(".xls", profile="rc4-cryptoapi", key_bits=128)

    def test_rc4_cryptoapi_40(self):
        _round_trip(".xls", profile="rc4-cryptoapi", key_bits=40)

    def test_rc4_legacy(self):
        _round_trip(".xls", profile="rc4-legacy")

    def test_xor(self):
        _round_trip(".xls", profile="xor")


class TestErrorPaths:
    def test_wrong_password_raises(self):
        import duke_sheets

        wb = _build_sample()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            wb.save_with_password(path, PASSWORD)
            with pytest.raises(Exception):
                duke_sheets.Workbook.open_with_password(path, "wrong-password")
        finally:
            os.unlink(path)

    def test_unknown_profile_raises(self):
        import duke_sheets

        wb = _build_sample()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            with pytest.raises(ValueError):
                wb.save_with_password(path, PASSWORD, profile="not-a-thing")
        finally:
            os.unlink(path)
