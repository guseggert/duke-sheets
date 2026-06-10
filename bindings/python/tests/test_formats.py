"""Plain (unencrypted) XLS, XLSB, and XLSX save -> open round-trips.

The binding dispatches the writer on the file extension; XLS and XLSB
carry binary formula token streams (BIFF8 / BIFF12), so formula text
and named-range survival exercises the compilers end to end.
"""

import os

import pytest


@pytest.fixture
def temp_dir(tmp_path):
    return str(tmp_path)


def _build_sample():
    import duke_sheets

    wb = duke_sheets.Workbook()
    wb.define_name("MyRange", "Sheet1!$A$1:$A$3")
    sheet = wb.get_sheet(0)
    sheet.set_cell("A1", 1.0)
    sheet.set_cell("A2", 2.0)
    sheet.set_cell("A3", 3.0)
    sheet.set_cell("B1", "label")
    sheet.set_cell("B2", True)
    sheet.set_formula("C1", "=SUM(A1:A3)")
    sheet.set_formula("C2", "=IF(A1>0,A2,A3)")
    sheet.set_formula("C3", "=SUM(MyRange)")
    wb.calculate()
    return wb


def _round_trip(temp_dir, ext):
    import duke_sheets

    path = os.path.join(temp_dir, f"sample{ext}")
    _build_sample().save(path)
    assert os.path.getsize(path) > 0

    opened = duke_sheets.Workbook.open(path)
    sheet = opened.get_sheet(0)
    assert sheet.get_cell("A1").as_number() == 1.0
    assert sheet.get_cell("B1").as_text() == "label"
    assert sheet.get_cell("B2").as_boolean() is True
    assert sheet.get_formula_at(0, 2) == "=SUM(A1:A3)"
    assert sheet.get_formula_at(1, 2) == "=IF(A1>0,A2,A3)"
    assert sheet.get_formula_at(2, 2) == "=SUM(MyRange)"
    assert sheet.get_cell("C1").as_number() == 6.0


class TestFormatRoundTrips:
    def test_xls_round_trip(self, temp_dir):
        _round_trip(temp_dir, ".xls")

    def test_xlsb_round_trip(self, temp_dir):
        _round_trip(temp_dir, ".xlsb")

    def test_xlsx_round_trip(self, temp_dir):
        _round_trip(temp_dir, ".xlsx")
