use ssfmt::ast::{Condition, DatePart, DigitPlaceholder, FormatPart, NamedColor, Section};
use ssfmt::NumberFormat;

#[test]
fn test_named_color_from_str() {
    assert_eq!("Red".parse::<NamedColor>().unwrap(), NamedColor::Red);
    assert_eq!("blue".parse::<NamedColor>().unwrap(), NamedColor::Blue);
    assert!("invalid".parse::<NamedColor>().is_err());
}

#[test]
fn test_condition_evaluate() {
    let cond = Condition::GreaterThan(100.0);
    assert!(cond.evaluate(150.0));
    assert!(!cond.evaluate(50.0));
    assert!(!cond.evaluate(100.0));
}

#[test]
fn test_digit_placeholder_properties() {
    assert!(DigitPlaceholder::Zero.is_required());
    assert!(!DigitPlaceholder::Hash.is_required());
    assert!(!DigitPlaceholder::Question.is_required());
}

#[test]
fn test_format_part_is_date_part() {
    let year = FormatPart::DatePart(DatePart::Year4);
    let digit = FormatPart::Digit(DigitPlaceholder::Zero);

    assert!(year.is_date_part());
    assert!(!digit.is_date_part());
}

#[test]
fn test_number_format_is_date_format() {
    // A format with date parts should be detected as date format
    let section = Section {
        condition: None,
        color: None,
        parts: vec![
            FormatPart::DatePart(DatePart::Year4),
            FormatPart::Literal("-".into()),
            FormatPart::DatePart(DatePart::Month2),
        ],
        metadata: ssfmt::ast::SectionMetadata::default(),
    };
    let format = NumberFormat::from_sections(vec![section]);
    assert!(format.is_date_format());
}

#[test]
fn test_number_format_sections_limit() {
    let sections: Vec<Section> = (0..5)
        .map(|_| Section {
            condition: None,
            color: None,
            parts: vec![],
            metadata: ssfmt::ast::SectionMetadata::default(),
        })
        .collect();
    // Should only keep first 4 sections
    let format = NumberFormat::from_sections(sections);
    assert_eq!(format.sections().len(), 4);
}
