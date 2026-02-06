//! Tests for formula evaluation with cell references

use duke_sheets::prelude::*;
use duke_sheets::{evaluate, parse_formula, EvaluationContext, FormulaValue};

/// Test basic formula evaluation without cell references
#[test]
fn test_evaluate_simple_formulas() {
    // Arithmetic
    let ast = parse_formula("=1+2*3").unwrap();
    let ctx = EvaluationContext::simple();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(7.0));

    // String concatenation
    let ast = parse_formula("=\"Hello \"&\"World\"").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::String("Hello World".into()));

    // Comparison
    let ast = parse_formula("=5>3").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));
}

/// Test SUM function
#[test]
fn test_evaluate_sum() {
    let ast = parse_formula("=SUM(1,2,3,4,5)").unwrap();
    let ctx = EvaluationContext::simple();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(15.0));
}

/// Test IF function
#[test]
fn test_evaluate_if() {
    let ctx = EvaluationContext::simple();

    let ast = parse_formula("=IF(1>0,\"Yes\",\"No\")").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::String("Yes".into()));

    let ast = parse_formula("=IF(1<0,\"Yes\",\"No\")").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::String("No".into()));
}

/// Test formula evaluation with cell references
#[test]
fn test_evaluate_with_cell_references() {
    // Create a workbook with some data
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 10.0).unwrap();
    sheet.set_cell_value("A2", 20.0).unwrap();
    sheet.set_cell_value("A3", 30.0).unwrap();
    sheet.set_cell_value("B1", 5.0).unwrap();

    // Create evaluation context with workbook reference
    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // Test simple cell reference
    let ast = parse_formula("=A1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(10.0));

    // Test cell reference in arithmetic
    let ast = parse_formula("=A1+B1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(15.0));

    // Test cell reference in comparison
    let ast = parse_formula("=A1>B1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));
}

/// Test formula evaluation with range references
#[test]
fn test_evaluate_with_range_references() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 10.0).unwrap();
    sheet.set_cell_value("A2", 20.0).unwrap();
    sheet.set_cell_value("A3", 30.0).unwrap();

    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // Test SUM with range
    let ast = parse_formula("=SUM(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(60.0));

    // Test AVERAGE with range
    let ast = parse_formula("=AVERAGE(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(20.0));

    // Test MIN/MAX with range
    let ast = parse_formula("=MIN(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(10.0));

    let ast = parse_formula("=MAX(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(30.0));
}

/// Test complex nested formulas
#[test]
fn test_evaluate_complex_formulas() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 100.0).unwrap();
    sheet.set_cell_value("A2", 50.0).unwrap();
    sheet.set_cell_value("B1", 0.1).unwrap(); // 10%

    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // Calculate: IF A1 > A2, calculate 10% of A1, else calculate 10% of A2
    let ast = parse_formula("=IF(A1>A2,A1*B1,A2*B1)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(10.0));

    // Nested SUM and multiplication
    let ast = parse_formula("=SUM(A1,A2)*B1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(15.0));
}

/// Test error propagation in formulas
#[test]
fn test_error_propagation() {
    let ctx = EvaluationContext::simple();

    // Division by zero
    let ast = parse_formula("=1/0").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert!(matches!(result, FormulaValue::Error(_)));

    // Error in arithmetic propagates
    let ast = parse_formula("=1/0+5").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert!(matches!(result, FormulaValue::Error(_)));
}

/// Test empty cell handling
#[test]
fn test_empty_cell_handling() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 10.0).unwrap();
    // A2 is empty
    sheet.set_cell_value("A3", 30.0).unwrap();

    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // Empty cells are treated as 0 in arithmetic
    let ast = parse_formula("=A1+A2").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(10.0)); // 10 + 0

    // SUM ignores empty cells
    let ast = parse_formula("=SUM(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(40.0)); // 10 + 0 + 30
}

/// Test string operations
#[test]
fn test_string_operations() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", "Hello").unwrap();
    sheet.set_cell_value("B1", "World").unwrap();

    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // String concatenation with cell references
    let ast = parse_formula("=A1&\" \"&B1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::String("Hello World".into()));
}

/// Test boolean functions
#[test]
fn test_boolean_functions() {
    let ctx = EvaluationContext::simple();

    // AND function
    let ast = parse_formula("=AND(TRUE,TRUE,TRUE)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));

    let ast = parse_formula("=AND(TRUE,FALSE,TRUE)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(false));

    // OR function
    let ast = parse_formula("=OR(FALSE,FALSE,TRUE)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));

    // NOT function
    let ast = parse_formula("=NOT(FALSE)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));

    // Combined
    let ast = parse_formula("=AND(NOT(FALSE),OR(TRUE,FALSE))").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));
}
