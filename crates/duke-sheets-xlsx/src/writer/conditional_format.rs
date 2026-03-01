use quick_xml::escape::escape;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use crate::styles::XlsxStyleTable;

use super::{write_color_element, XlsxResult, XmlWriter};

pub(super) fn write_conditional_formatting(
    w: &mut XmlWriter,
    sheet: &duke_sheets_core::Worksheet,
    sheet_index: usize,
    style_table: &XlsxStyleTable,
) -> XlsxResult<()> {
    use duke_sheets_core::conditional_format::CfRuleType;

    let rules = sheet.conditional_formats();
    if rules.is_empty() {
        return Ok(());
    }

    for (rule_idx, rule) in rules.iter().enumerate() {
        if rule.ranges.is_empty() {
            continue;
        }

        let sqref: String = rule
            .ranges
            .iter()
            .map(|r| r.to_string())
            .collect::<Vec<_>>()
            .join(" ");

        let mut cf_tag = BytesStart::new("conditionalFormatting");
        cf_tag.push_attribute(("sqref", sqref.as_str()));
        w.write_event(Event::Start(cf_tag))?;

        let rule_type = rule.rule_type.xlsx_type();
        let dxf_id = style_table
            .dxf_id_for(sheet_index, rule_idx)
            .or(rule.dxf_id);
        let priority_val = rule.priority.max(1);
        let priority_s = priority_val.to_string();

        match &rule.rule_type {
            CfRuleType::CellIs {
                operator,
                formula1,
                formula2,
            } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("operator", operator.xlsx_operator()));
                tag.push_attribute(("priority", priority_s.as_str()));
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Start(tag))?;

                w.create_element("formula")
                    .write_text_content(BytesText::new(formula1))?;
                if let Some(f2) = formula2 {
                    w.create_element("formula")
                        .write_text_content(BytesText::new(f2))?;
                }
                w.write_event(Event::End(BytesEnd::new("cfRule")))?;
            }

            CfRuleType::Expression { formula } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Start(tag))?;

                w.create_element("formula")
                    .write_text_content(BytesText::new(formula))?;
                w.write_event(Event::End(BytesEnd::new("cfRule")))?;
            }

            CfRuleType::ColorScale { colors } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                if rule.stop_if_true {
                    tag.push_attribute(("stopIfTrue", "1"));
                }
                w.write_event(Event::Start(tag))?;

                w.write_event(Event::Start(BytesStart::new("colorScale")))?;
                for cv in colors {
                    let mut cfvo = BytesStart::new("cfvo");
                    cfvo.push_attribute(("type", cv.value_type.xlsx_type()));
                    if let Some(ref v) = cv.value {
                        cfvo.push_attribute(("val", v.as_str()));
                    }
                    w.write_event(Event::Empty(cfvo))?;
                }
                for cv in colors {
                    write_color_element(w, "color", &cv.color)?;
                }
                w.write_event(Event::End(BytesEnd::new("colorScale")))?;
                w.write_event(Event::End(BytesEnd::new("cfRule")))?;
            }

            CfRuleType::DataBar {
                min_value,
                max_value,
                color,
                show_value,
                ..
            } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                if rule.stop_if_true {
                    tag.push_attribute(("stopIfTrue", "1"));
                }
                w.write_event(Event::Start(tag))?;

                let mut db = BytesStart::new("dataBar");
                if !*show_value {
                    db.push_attribute(("showValue", "0"));
                }
                w.write_event(Event::Start(db))?;

                let mut cfvo_min = BytesStart::new("cfvo");
                cfvo_min.push_attribute(("type", min_value.value_type.xlsx_type()));
                if let Some(ref v) = min_value.value {
                    cfvo_min.push_attribute(("val", v.as_str()));
                }
                w.write_event(Event::Empty(cfvo_min))?;

                let mut cfvo_max = BytesStart::new("cfvo");
                cfvo_max.push_attribute(("type", max_value.value_type.xlsx_type()));
                if let Some(ref v) = max_value.value {
                    cfvo_max.push_attribute(("val", v.as_str()));
                }
                w.write_event(Event::Empty(cfvo_max))?;

                write_color_element(w, "color", color)?;

                w.write_event(Event::End(BytesEnd::new("dataBar")))?;
                w.write_event(Event::End(BytesEnd::new("cfRule")))?;
            }

            CfRuleType::IconSet {
                icon_style,
                values,
                reverse,
                show_value,
            } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                if rule.stop_if_true {
                    tag.push_attribute(("stopIfTrue", "1"));
                }
                w.write_event(Event::Start(tag))?;

                let mut is_tag = BytesStart::new("iconSet");
                is_tag.push_attribute(("iconSet", icon_style.xlsx_name()));
                if *reverse {
                    is_tag.push_attribute(("reverse", "1"));
                }
                if !*show_value {
                    is_tag.push_attribute(("showValue", "0"));
                }
                w.write_event(Event::Start(is_tag))?;

                for val in values {
                    let mut cfvo = BytesStart::new("cfvo");
                    cfvo.push_attribute(("type", val.value_type.xlsx_type()));
                    if let Some(ref v) = val.value {
                        cfvo.push_attribute(("val", v.as_str()));
                    }
                    w.write_event(Event::Empty(cfvo))?;
                }

                w.write_event(Event::End(BytesEnd::new("iconSet")))?;
                w.write_event(Event::End(BytesEnd::new("cfRule")))?;
            }

            CfRuleType::Top10 {
                rank,
                percent,
                bottom,
            } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                let rank_s = rank.to_string();
                tag.push_attribute(("rank", rank_s.as_str()));
                if *percent {
                    tag.push_attribute(("percent", "1"));
                }
                if *bottom {
                    tag.push_attribute(("bottom", "1"));
                }
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Empty(tag))?;
            }

            CfRuleType::AboveAverage {
                above,
                equal_average,
                std_dev,
            } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                if !*above {
                    tag.push_attribute(("aboveAverage", "0"));
                }
                if *equal_average {
                    tag.push_attribute(("equalAverage", "1"));
                }
                if let Some(s) = std_dev {
                    let v = s.to_string();
                    tag.push_attribute(("stdDev", v.as_str()));
                }
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Empty(tag))?;
            }

            CfRuleType::ContainsText { text } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                let text_esc = escape(text.as_str());
                tag.push_attribute(("text", &*text_esc));
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Start(tag))?;

                let first_cell = sqref.split(' ').next().unwrap_or("A1");
                let formula = format!(
                    "NOT(ISERROR(SEARCH(\"{}\",{})))",
                    text.replace('"', "\"\""),
                    first_cell
                );
                w.create_element("formula")
                    .write_text_content(BytesText::new(&formula))?;
                w.write_event(Event::End(BytesEnd::new("cfRule")))?;
            }

            CfRuleType::BeginsWith { text } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                let text_esc = escape(text.as_str());
                tag.push_attribute(("text", &*text_esc));
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Start(tag))?;

                let first_cell = sqref
                    .split(' ')
                    .next()
                    .unwrap_or("A1")
                    .split(':')
                    .next()
                    .unwrap_or("A1");
                let formula = format!(
                    "LEFT({},{})=\"{}\"",
                    first_cell,
                    text.len(),
                    text.replace('"', "\"\"")
                );
                w.create_element("formula")
                    .write_text_content(BytesText::new(&formula))?;
                w.write_event(Event::End(BytesEnd::new("cfRule")))?;
            }

            CfRuleType::EndsWith { text } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                let text_esc = escape(text.as_str());
                tag.push_attribute(("text", &*text_esc));
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Start(tag))?;

                let first_cell = sqref
                    .split(' ')
                    .next()
                    .unwrap_or("A1")
                    .split(':')
                    .next()
                    .unwrap_or("A1");
                let formula = format!(
                    "RIGHT({},{})=\"{}\"",
                    first_cell,
                    text.len(),
                    text.replace('"', "\"\"")
                );
                w.create_element("formula")
                    .write_text_content(BytesText::new(&formula))?;
                w.write_event(Event::End(BytesEnd::new("cfRule")))?;
            }

            CfRuleType::DuplicateValues
            | CfRuleType::UniqueValues
            | CfRuleType::ContainsBlanks
            | CfRuleType::NotContainsBlanks
            | CfRuleType::ContainsErrors
            | CfRuleType::NotContainsErrors => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Empty(tag))?;
            }

            CfRuleType::TimePeriod { period } => {
                let mut tag = BytesStart::new("cfRule");
                tag.push_attribute(("type", rule_type));
                tag.push_attribute(("priority", priority_s.as_str()));
                tag.push_attribute(("timePeriod", period.xlsx_period()));
                push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                w.write_event(Event::Empty(tag))?;
            }
        }

        w.write_event(Event::End(BytesEnd::new("conditionalFormatting")))?;
    }

    Ok(())
}

pub(super) fn push_dxf_and_stop(tag: &mut BytesStart, dxf_id: Option<u32>, stop_if_true: bool) {
    if let Some(id) = dxf_id {
        let s = id.to_string();
        tag.push_attribute(("dxfId", s.as_str()));
    }
    if stop_if_true {
        tag.push_attribute(("stopIfTrue", "1"));
    }
}
