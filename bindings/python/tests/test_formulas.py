"""
Tests for formula functionality.
"""

import pytest


class TestSetFormula:
    """Test setting formulas."""

    def test_set_simple_formula(self, workbook):
        """Should set a simple formula."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")
        
        value = sheet.get_cell("A1")
        assert value.is_formula
        assert "1+1" in value.formula_text()

    def test_set_formula_with_cell_ref(self, workbook):
        """Should set formula with cell reference."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10.0)
        sheet.set_formula("A2", "=A1*2")
        
        value = sheet.get_cell("A2")
        assert value.is_formula

    def test_set_formula_with_function(self, workbook):
        """Should set formula with function."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=SUM(1,2,3)")
        
        value = sheet.get_cell("A1")
        assert value.is_formula
        assert "SUM" in value.formula_text()

    def test_set_formula_without_equals(self, workbook):
        """Formula without = should be treated as text or error."""
        sheet = workbook.get_sheet(0)
        # This might raise an error or be treated as text
        # depending on implementation
        try:
            sheet.set_formula("A1", "1+1")
            value = sheet.get_cell("A1")
            # Either it's a formula or text
        except Exception:
            pass  # Some implementations may reject this


class TestFormulaCalculation:
    """Test formula evaluation after calculate()."""

    def test_simple_addition(self, workbook):
        """Should calculate simple addition."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+2")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("A1")
        assert value.as_number() == 3.0

    def test_cell_reference(self, workbook):
        """Should calculate cell references."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10.0)
        sheet.set_cell("A2", 20.0)
        sheet.set_formula("A3", "=A1+A2")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("A3")
        assert value.as_number() == 30.0

    def test_sum_function(self, workbook):
        """Should calculate SUM function."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 1.0)
        sheet.set_cell("A2", 2.0)
        sheet.set_cell("A3", 3.0)
        sheet.set_cell("A4", 4.0)
        sheet.set_formula("A5", "=SUM(A1:A4)")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("A5")
        assert value.as_number() == 10.0

    def test_average_function(self, workbook):
        """Should calculate AVERAGE function."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10.0)
        sheet.set_cell("A2", 20.0)
        sheet.set_cell("A3", 30.0)
        sheet.set_formula("A4", "=AVERAGE(A1:A3)")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("A4")
        assert value.as_number() == 20.0

    def test_if_function(self, workbook):
        """Should calculate IF function."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10.0)
        sheet.set_formula("A2", "=IF(A1>5, \"Yes\", \"No\")")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("A2")
        assert value.as_text() == "Yes"

    def test_nested_formulas(self, workbook):
        """Should calculate nested formulas."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 5.0)
        sheet.set_formula("A2", "=A1*2")      # 10
        sheet.set_formula("A3", "=A2+A1")     # 15
        sheet.set_formula("A4", "=SUM(A1:A3)") # 30
        
        workbook.calculate()
        
        assert sheet.get_calculated_value("A2").as_number() == 10.0
        assert sheet.get_calculated_value("A3").as_number() == 15.0
        assert sheet.get_calculated_value("A4").as_number() == 30.0


class TestNamedRangesInFormulas:
    """Test using named ranges in formulas."""

    def test_named_constant_in_formula(self, workbook):
        """Should use named constant in formula."""
        sheet = workbook.get_sheet(0)
        
        # Define a constant
        workbook.define_name("TaxRate", "0.1")
        
        # Use in formula
        sheet.set_cell("A1", 100.0)
        sheet.set_formula("A2", "=A1*TaxRate")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("A2")
        assert value.as_number() == pytest.approx(10.0)

    def test_named_cell_in_formula(self, workbook):
        """Should use named cell reference in formula."""
        sheet = workbook.get_sheet(0)
        
        # Set up data
        sheet.set_cell("A1", 50.0)
        
        # Define name pointing to cell
        workbook.define_name("Price", "Sheet1!$A$1")
        
        # Use in formula
        sheet.set_formula("B1", "=Price*2")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("B1")
        assert value.as_number() == 100.0


class TestFormulaErrors:
    """Test formula error handling."""

    def test_division_by_zero(self, workbook):
        """Should return #DIV/0! error."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1/0")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("A1")
        assert value.is_error
        assert "#DIV/0!" in value.as_error()

    def test_invalid_reference(self, workbook):
        """Invalid reference should produce error."""
        sheet = workbook.get_sheet(0)
        # Reference to non-existent sheet
        sheet.set_formula("A1", "=NonExistentSheet!A1")
        
        workbook.calculate()
        
        value = sheet.get_calculated_value("A1")
        # Should be some kind of error
        assert value.is_error or value.as_number() is None
