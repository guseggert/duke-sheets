use duke_sheets_core::conditional_format::{
    CfOperator, CfRuleType, ConditionalFormatRule, TimePeriod,
};
use duke_sheets_core::style::Color;
use duke_sheets_core::CellRange;

use super::ThemePalette;

pub(crate) fn parse_color_element(
    e: &quick_xml::events::BytesStart,
    theme_palette: Option<&ThemePalette>,
) -> Color {
    let mut rgb: Option<String> = None;
    let mut theme: Option<u8> = None;
    let mut tint: Option<f64> = None;
    let mut indexed: Option<u8> = None;
    let mut auto = false;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"rgb" => {
                rgb = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"theme" => {
                theme = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<u8>().ok());
            }
            b"tint" => {
                tint = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<f64>().ok());
            }
            b"indexed" => {
                indexed = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<u8>().ok());
            }
            b"auto" => {
                auto = attr.unescape_value().ok().as_deref() == Some("1");
            }
            _ => {}
        }
    }

    if let Some(rgb_str) = rgb {
        let hex = rgb_str.trim_start_matches('#');

        if hex.len() == 8 {
            if let (Ok(_a), Ok(r), Ok(g), Ok(b)) = (
                u8::from_str_radix(&hex[0..2], 16),
                u8::from_str_radix(&hex[2..4], 16),
                u8::from_str_radix(&hex[4..6], 16),
                u8::from_str_radix(&hex[6..8], 16),
            ) {
                return Color::Rgb { r, g, b };
            }
        } else if hex.len() == 6 {
            if let (Ok(r), Ok(g), Ok(b)) = (
                u8::from_str_radix(&hex[0..2], 16),
                u8::from_str_radix(&hex[2..4], 16),
                u8::from_str_radix(&hex[4..6], 16),
            ) {
                return Color::Rgb { r, g, b };
            }
        }
    }

    if let Some(index) = theme {
        let tint_i8 = tint.map(|t| (t * 100.0).round() as i8).unwrap_or(0);
        if let Some(theme) = theme_palette {
            let (r, g, b) = theme.resolve_theme(index, tint_i8);
            return Color::Rgb { r, g, b };
        }
        return Color::Theme {
            index,
            tint: tint_i8,
        };
    }

    if let Some(i) = indexed {
        return Color::Indexed(i);
    }

    if auto {
        return Color::Auto;
    }

    Color::Auto
}

pub(super) fn parse_cf_rule_attrs(
    e: &quick_xml::events::BytesStart,
    sqref: Option<&str>,
) -> ConditionalFormatRule {
    let mut rule = ConditionalFormatRule::default();
    let mut rule_type: Option<String> = None;
    let mut operator: Option<String> = None;
    let mut text: Option<String> = None;
    let mut rank: Option<u32> = None;
    let mut percent = false;
    let mut bottom = false;
    let mut above_average = true;
    let mut equal_average = false;
    let mut std_dev: Option<u32> = None;
    let mut time_period: Option<String> = None;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"type" => {
                rule_type = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"operator" => {
                operator = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"priority" => {
                if let Some(p) = attr.unescape_value().ok().and_then(|s| s.parse().ok()) {
                    rule.priority = p;
                }
            }
            b"stopIfTrue" => {
                rule.stop_if_true = attr.unescape_value().ok().is_some_and(|s| s == "1");
            }
            b"dxfId" => {
                rule.dxf_id = attr.unescape_value().ok().and_then(|s| s.parse().ok());
            }
            b"text" => {
                text = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"rank" => {
                rank = attr.unescape_value().ok().and_then(|s| s.parse().ok());
            }
            b"percent" => {
                percent = attr.unescape_value().ok().is_some_and(|s| s == "1");
            }
            b"bottom" => {
                bottom = attr.unescape_value().ok().is_some_and(|s| s == "1");
            }
            b"aboveAverage" => {
                above_average = attr.unescape_value().ok().is_none_or(|s| s != "0");
            }
            b"equalAverage" => {
                equal_average = attr.unescape_value().ok().is_some_and(|s| s == "1");
            }
            b"stdDev" => {
                std_dev = attr.unescape_value().ok().and_then(|s| s.parse().ok());
            }
            b"timePeriod" => {
                time_period = attr.unescape_value().ok().map(|s| s.to_string());
            }
            _ => {}
        }
    }

    if let Some(sqref) = sqref {
        rule.ranges = parse_sqref(sqref);
    }

    let op = operator
        .as_deref()
        .and_then(CfOperator::from_xlsx)
        .unwrap_or(CfOperator::Equal);

    rule.rule_type = match rule_type.as_deref() {
        Some("cellIs") => CfRuleType::CellIs {
            operator: op,
            formula1: String::new(),
            formula2: None,
        },
        Some("expression") => CfRuleType::Expression {
            formula: String::new(),
        },
        Some("top10") => CfRuleType::Top10 {
            rank: rank.unwrap_or(10),
            percent,
            bottom,
        },
        Some("aboveAverage") => CfRuleType::AboveAverage {
            above: above_average,
            equal_average,
            std_dev,
        },
        Some("containsText") => CfRuleType::ContainsText {
            text: text.unwrap_or_default(),
        },
        Some("beginsWith") => CfRuleType::BeginsWith {
            text: text.unwrap_or_default(),
        },
        Some("endsWith") => CfRuleType::EndsWith {
            text: text.unwrap_or_default(),
        },
        Some("duplicateValues") => CfRuleType::DuplicateValues,
        Some("uniqueValues") => CfRuleType::UniqueValues,
        Some("containsBlanks") => CfRuleType::ContainsBlanks,
        Some("notContainsBlanks") => CfRuleType::NotContainsBlanks,
        Some("containsErrors") => CfRuleType::ContainsErrors,
        Some("notContainsErrors") => CfRuleType::NotContainsErrors,
        Some("timePeriod") => CfRuleType::TimePeriod {
            period: time_period
                .as_deref()
                .and_then(TimePeriod::from_xlsx)
                .unwrap_or(TimePeriod::Today),
        },
        _ => CfRuleType::Expression {
            formula: String::new(),
        },
    };

    rule
}

pub(super) fn apply_cf_formulas(rule: &mut ConditionalFormatRule, formulas: &[String]) {
    match &mut rule.rule_type {
        CfRuleType::CellIs {
            formula1, formula2, ..
        } => {
            if let Some(f1) = formulas.first() {
                *formula1 = f1.clone();
            }
            *formula2 = formulas.get(1).cloned();
        }
        CfRuleType::Expression { formula } => {
            if let Some(f1) = formulas.first() {
                *formula = f1.clone();
            }
        }
        _ => {}
    }
}

pub(super) fn parse_sqref(sqref: &str) -> Vec<CellRange> {
    sqref
        .split_whitespace()
        .filter_map(|s| CellRange::parse(s).ok())
        .collect()
}

#[cfg(test)]
mod tests {
    use quick_xml::events::BytesStart;

    use super::*;

    #[test]
    fn test_parse_color_element_theme_and_tint() {
        let mut e = BytesStart::new("color");
        e.push_attribute(("theme", "4"));
        e.push_attribute(("tint", "0.5"));

        assert_eq!(
            parse_color_element(&e, None),
            Color::Theme { index: 4, tint: 50 }
        );
    }

    #[test]
    fn test_parse_color_element_indexed_and_auto() {
        let mut indexed = BytesStart::new("color");
        indexed.push_attribute(("indexed", "12"));
        assert_eq!(parse_color_element(&indexed, None), Color::Indexed(12));

        let mut auto = BytesStart::new("color");
        auto.push_attribute(("auto", "1"));
        assert_eq!(parse_color_element(&auto, None), Color::Auto);
    }
}
