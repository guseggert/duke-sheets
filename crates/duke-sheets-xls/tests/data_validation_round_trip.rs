//! Round-trip tests for the XLS writer's DVAL (0x01B2) and DV
//! (0x01BE) records covering the major validation types: list,
//! whole-number, decimal, text-length, custom formula, plus input
//! messages and error alerts.

use std::io::Cursor;

use duke_sheets_core::validation::{
    DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
};
use duke_sheets_core::{CellAddress, CellRange, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn range(start: &str, end: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start).expect("start"),
        CellAddress::parse(end).expect("end"),
    )
}

#[test]
fn list_inline_values_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut v = DataValidation::list("Red,Green,Blue");
    v.ranges = vec![range("A1", "A10")];
    ws.add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    assert_eq!(validations.len(), 1);
    match &validations[0].validation_type {
        ValidationType::List { source } => {
            assert_eq!(source, "Red,Green,Blue");
        }
        other => panic!("expected List, got {other:?}"),
    }
    assert_eq!(validations[0].ranges.len(), 1);
    assert_eq!(
        validations[0].ranges[0].start,
        CellAddress::parse("A1").unwrap()
    );
    assert_eq!(
        validations[0].ranges[0].end,
        CellAddress::parse("A10").unwrap()
    );
}

#[test]
fn list_cell_range_source_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Lookup").expect("rename");
    wb.add_worksheet_with_name("Form").expect("add");
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "Yes")
        .expect("A1");
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A2", "No")
        .expect("A2");

    let mut v = DataValidation::list("=Lookup!$A$1:$A$2");
    v.ranges = vec![range("B1", "B5")];
    wb.worksheet_mut(1).unwrap().add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet_by_name("Form").unwrap().data_validations();
    assert_eq!(validations.len(), 1);
    match &validations[0].validation_type {
        ValidationType::List { source } => {
            assert!(source.contains("Lookup"), "got {source:?}");
            // Reader emits absolute refs as $A$1; tolerate both styles.
            assert!(
                source.contains("A1") || source.contains("$A$1"),
                "got {source:?}"
            );
            assert!(
                source.contains("A2") || source.contains("$A$2"),
                "got {source:?}"
            );
        }
        other => panic!("expected List with cell-range source, got {other:?}"),
    }
}

#[test]
fn whole_number_between_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut v = DataValidation::default();
    v.validation_type = ValidationType::Whole {
        operator: ValidationOperator::Between,
        value1: "1".into(),
        value2: Some("100".into()),
    };
    v.ranges = vec![range("B1", "B10")];
    ws.add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    assert_eq!(validations.len(), 1);
    match &validations[0].validation_type {
        ValidationType::Whole {
            operator,
            value1,
            value2,
        } => {
            assert_eq!(*operator, ValidationOperator::Between);
            assert_eq!(value1, "1");
            assert_eq!(value2.as_deref(), Some("100"));
        }
        other => panic!("expected Whole, got {other:?}"),
    }
}

#[test]
fn whole_number_greater_than_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut v = DataValidation::default();
    v.validation_type = ValidationType::Whole {
        operator: ValidationOperator::GreaterThan,
        value1: "0".into(),
        value2: None,
    };
    v.ranges = vec![range("C1", "C5")];
    ws.add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    match &validations[0].validation_type {
        ValidationType::Whole {
            operator, value1, ..
        } => {
            assert_eq!(*operator, ValidationOperator::GreaterThan);
            assert_eq!(value1, "0");
        }
        other => panic!("expected Whole>, got {other:?}"),
    }
}

#[test]
fn decimal_validation_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut v = DataValidation::default();
    v.validation_type = ValidationType::Decimal {
        operator: ValidationOperator::Between,
        value1: "0.0".into(),
        value2: Some("100.5".into()),
    };
    v.ranges = vec![range("E1", "E5")];
    ws.add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    match &validations[0].validation_type {
        ValidationType::Decimal {
            operator,
            value1,
            value2,
        } => {
            assert_eq!(*operator, ValidationOperator::Between);
            assert!(
                value1.starts_with("0") && (value1.contains("0.0") || *value1 == "0"),
                "got {value1:?}"
            );
            assert!(
                value2.as_ref().is_some_and(|s| s.starts_with("100")),
                "got {value2:?}"
            );
        }
        other => panic!("expected Decimal, got {other:?}"),
    }
}

#[test]
fn text_length_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut v = DataValidation::default();
    v.validation_type = ValidationType::TextLength {
        operator: ValidationOperator::LessThanOrEqual,
        value1: "20".into(),
        value2: None,
    };
    v.ranges = vec![range("D1", "D5")];
    ws.add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    match &validations[0].validation_type {
        ValidationType::TextLength {
            operator, value1, ..
        } => {
            assert_eq!(*operator, ValidationOperator::LessThanOrEqual);
            assert_eq!(value1, "20");
        }
        other => panic!("expected TextLength, got {other:?}"),
    }
}

#[test]
fn custom_formula_validation_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 10.0).expect("A1");
    let mut v = DataValidation::default();
    v.validation_type = ValidationType::Custom {
        formula: "=A1>0".into(),
    };
    v.ranges = vec![range("B1", "B5")];
    ws.add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    match &validations[0].validation_type {
        ValidationType::Custom { formula } => {
            // The reader normalises formulas; just confirm the
            // operator and reference made it through.
            assert!(formula.contains('>'), "got {formula:?}");
            assert!(formula.contains("A1"), "got {formula:?}");
        }
        other => panic!("expected Custom, got {other:?}"),
    }
}

#[test]
fn input_message_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut v = DataValidation::list("Yes,No");
    v.ranges = vec![range("A1", "A1")];
    v.show_input_message = true;
    v.input_title = Some("Pick one".into());
    v.input_message = Some("Enter Yes or No".into());
    ws.add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    assert_eq!(validations.len(), 1);
    let dv = &validations[0];
    assert!(dv.show_input_message);
    assert_eq!(dv.input_title.as_deref(), Some("Pick one"));
    assert_eq!(dv.input_message.as_deref(), Some("Enter Yes or No"));
}

#[test]
fn error_alert_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut v = DataValidation::default();
    v.validation_type = ValidationType::Whole {
        operator: ValidationOperator::Between,
        value1: "1".into(),
        value2: Some("10".into()),
    };
    v.ranges = vec![range("A1", "A1")];
    v.show_error_alert = true;
    v.error_style = ValidationErrorStyle::Warning;
    v.error_title = Some("Out of range".into());
    v.error_message = Some("Please enter 1-10.".into());
    ws.add_data_validation(v);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    let dv = &validations[0];
    assert!(dv.show_error_alert);
    assert_eq!(dv.error_style, ValidationErrorStyle::Warning);
    assert_eq!(dv.error_title.as_deref(), Some("Out of range"));
    assert_eq!(dv.error_message.as_deref(), Some("Please enter 1-10."));
}

#[test]
fn multiple_validations_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut v1 = DataValidation::list("A,B,C");
    v1.ranges = vec![range("A1", "A5")];
    let mut v2 = DataValidation::default();
    v2.validation_type = ValidationType::Whole {
        operator: ValidationOperator::Equal,
        value1: "42".into(),
        value2: None,
    };
    v2.ranges = vec![range("B1", "B5")];
    let mut v3 = DataValidation::default();
    v3.validation_type = ValidationType::TextLength {
        operator: ValidationOperator::GreaterThan,
        value1: "3".into(),
        value2: None,
    };
    v3.ranges = vec![range("C1", "C5")];

    ws.add_data_validation(v1);
    ws.add_data_validation(v2);
    ws.add_data_validation(v3);

    let parsed = write_then_read(&wb);
    let validations = parsed.worksheet(0).unwrap().data_validations();
    assert_eq!(validations.len(), 3);
}

#[test]
fn no_validations_emits_no_dval_record() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("A1");

    let parsed = write_then_read(&wb);
    assert!(parsed.worksheet(0).unwrap().data_validations().is_empty());
}
