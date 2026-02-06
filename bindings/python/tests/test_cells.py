"""
Tests for CellValue class and cell types.
"""

import pytest


class TestCellValueTypes:
    """Test CellValue type checking."""

    def test_empty_value(self, workbook):
        """Empty cell should have is_empty True."""
        sheet = workbook.get_sheet(0)
        value = sheet.get_cell("Z99")
        
        assert value.is_empty
        assert not value.is_number
        assert not value.is_text
        assert not value.is_boolean
        assert not value.is_error
        assert not value.is_formula

    def test_number_value(self, workbook):
        """Number cell should have is_number True."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.5)
        value = sheet.get_cell("A1")
        
        assert value.is_number
        assert not value.is_empty
        assert not value.is_text
        assert not value.is_boolean

    def test_text_value(self, workbook):
        """Text cell should have is_text True."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "Hello")
        value = sheet.get_cell("A1")
        
        assert value.is_text
        assert not value.is_empty
        assert not value.is_number
        assert not value.is_boolean

    def test_boolean_value(self, workbook):
        """Boolean cell should have is_boolean True."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", True)
        value = sheet.get_cell("A1")
        
        assert value.is_boolean
        assert not value.is_empty
        assert not value.is_number
        assert not value.is_text

    def test_formula_value(self, workbook):
        """Formula cell should have is_formula True."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")
        value = sheet.get_cell("A1")
        
        assert value.is_formula
        assert not value.is_empty


class TestCellValueConversions:
    """Test CellValue type conversion methods."""

    def test_as_number_from_number(self, workbook):
        """as_number should return float for number cell."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 3.14159)
        value = sheet.get_cell("A1")
        
        assert value.as_number() == pytest.approx(3.14159)

    def test_as_number_from_text(self, workbook):
        """as_number should return None for text cell."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "Hello")
        value = sheet.get_cell("A1")
        
        assert value.as_number() is None

    def test_as_text_from_text(self, workbook):
        """as_text should return string for text cell."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "Hello")
        value = sheet.get_cell("A1")
        
        assert value.as_text() == "Hello"

    def test_as_text_from_number(self, workbook):
        """as_text should return None for number cell."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.0)
        value = sheet.get_cell("A1")
        
        assert value.as_text() is None

    def test_as_boolean_from_boolean(self, workbook):
        """as_boolean should return bool for boolean cell."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", True)
        value = sheet.get_cell("A1")
        
        assert value.as_boolean() == True

    def test_as_boolean_from_number(self, workbook):
        """as_boolean should return None for number cell."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.0)
        value = sheet.get_cell("A1")
        
        assert value.as_boolean() is None

    def test_formula_text(self, workbook):
        """formula_text should return formula string."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=SUM(B1:B10)")
        value = sheet.get_cell("A1")
        
        formula = value.formula_text()
        assert formula is not None
        assert "SUM" in formula

    def test_formula_text_from_non_formula(self, workbook):
        """formula_text should return None for non-formula."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.0)
        value = sheet.get_cell("A1")
        
        assert value.formula_text() is None


class TestCellValueToPython:
    """Test CellValue.to_python() conversion."""

    def test_empty_to_python(self, workbook):
        """Empty should convert to None."""
        sheet = workbook.get_sheet(0)
        value = sheet.get_cell("Z99")
        
        assert value.to_python() is None

    def test_number_to_python(self, workbook):
        """Number should convert to float."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.5)
        value = sheet.get_cell("A1")
        
        py_val = value.to_python()
        assert isinstance(py_val, float)
        assert py_val == 42.5

    def test_text_to_python(self, workbook):
        """Text should convert to str."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "Hello")
        value = sheet.get_cell("A1")
        
        py_val = value.to_python()
        assert isinstance(py_val, str)
        assert py_val == "Hello"

    def test_boolean_to_python(self, workbook):
        """Boolean should convert to bool."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", True)
        value = sheet.get_cell("A1")
        
        py_val = value.to_python()
        assert isinstance(py_val, bool)
        assert py_val == True


class TestCellValueRepr:
    """Test CellValue string representations."""

    def test_repr_empty(self, workbook):
        """Empty cell repr."""
        sheet = workbook.get_sheet(0)
        value = sheet.get_cell("Z99")
        
        r = repr(value)
        assert "CellValue" in r
        assert "Empty" in r

    def test_repr_number(self, workbook):
        """Number cell repr."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.0)
        value = sheet.get_cell("A1")
        
        r = repr(value)
        assert "CellValue" in r
        assert "Number" in r
        assert "42" in r

    def test_repr_text(self, workbook):
        """Text cell repr."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "Hello")
        value = sheet.get_cell("A1")
        
        r = repr(value)
        assert "CellValue" in r
        assert "Text" in r
        assert "Hello" in r

    def test_str_empty(self, workbook):
        """Empty cell str should be empty string."""
        sheet = workbook.get_sheet(0)
        value = sheet.get_cell("Z99")
        
        assert str(value) == ""

    def test_str_number(self, workbook):
        """Number cell str."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.0)
        value = sheet.get_cell("A1")
        
        assert str(value) == "42"

    def test_str_boolean_true(self, workbook):
        """Boolean True cell str."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", True)
        value = sheet.get_cell("A1")
        
        assert str(value) == "TRUE"

    def test_str_boolean_false(self, workbook):
        """Boolean False cell str."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", False)
        value = sheet.get_cell("A1")
        
        assert str(value) == "FALSE"
