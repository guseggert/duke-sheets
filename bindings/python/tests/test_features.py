"""Populated-feature reads via runtime-generated fixtures.

Comments, autofilters, data validations, and embedded images are
read-only in the binding, so fixtures carrying them are generated at
test time by the Rust fixture generator (binary fixtures are never
committed).
"""

import os
import subprocess

import pytest

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))


@pytest.fixture(scope="session")
def fixture_dir(tmp_path_factory):
    out = str(tmp_path_factory.mktemp("duke-fixtures"))
    subprocess.run(
        [
            "cargo",
            "run",
            "-p",
            "duke-sheets",
            "--features",
            "full",
            "--example",
            "gen_binding_fixtures",
            "--",
            out,
        ],
        cwd=REPO_ROOT,
        check=True,
    )
    return out


@pytest.fixture(params=["xlsx", "xls", "xlsb"])
def opened(request, fixture_dir):
    import duke_sheets

    ext = request.param
    wb = duke_sheets.Workbook.open(os.path.join(fixture_dir, f"sample.{ext}"))
    return ext, wb.get_sheet(0)


class TestPopulatedFeatureReads:
    def test_comment(self, opened):
        _, sheet = opened
        comment = sheet.get_comment("A1")
        assert comment is not None
        assert comment.author == "Tester"
        assert comment.text == "fixture comment"
        assert sheet.comment_count == 1

    def test_autofilter(self, opened):
        _, sheet = opened
        af = sheet.auto_filter
        assert af is not None
        assert af.range == "A1:A4"
        assert len(af.filter_columns) == 1
        col = af.filter_columns[0]
        assert col.col_id == 0
        assert col.filter_type == "values"
        assert col.values == ["1", "3"]

    def test_data_validation(self, opened):
        _, sheet = opened
        dvs = sheet.data_validations
        assert len(dvs) == 1
        assert dvs[0].validation_type == "list"
        assert dvs[0].list_source == "Red,Green,Blue"
        assert "C1:C5" in dvs[0].ranges

    def test_values_and_named_range_formula(self, opened):
        _, sheet = opened
        assert sheet.get_cell("A1").as_text() == "Score"
        assert sheet.get_formula_at(0, 1) == "=SUM(MyRange)"

    def test_embedded_image(self, opened):
        ext, sheet = opened
        if ext == "xlsb":
            pytest.skip("XLSB image reading not supported yet")
        assert sheet.image_count == 1
        drawing = sheet.images[0]
        assert drawing.name == "FixturePic"
        assert drawing.image.format.lower() == "png"
        data = sheet.drawing_image_data(drawing.drawing_path)
        assert len(data) > 0
        assert data[:2] == b"\x89P"
