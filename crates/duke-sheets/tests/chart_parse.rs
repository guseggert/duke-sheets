//! Standard (`c:`) chart parsing.

mod xml_forms;

use duke_sheets_chart::Chart;

fn parse(xml: &str) -> Chart {
    duke_sheets_chart::parse::parse_chart_xml(xml.as_bytes()).expect("parse chart")
}

/// Leading and trailing whitespace inside a text-bearing element is part
/// of the value: a series really can be named " Q1 " and a label
/// separator really can be " | ".
#[test]
fn whitespace_inside_text_elements_is_part_of_the_value() {
    let doc = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t> Padded Title </a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:tx><c:strRef><c:f> S!$B$1 </c:f></c:strRef></c:tx><c:dLbls><c:separator> | </c:separator></c:dLbls><c:val><c:numRef><c:f> S!$B$2:$B$4 </c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

    let chart = parse(doc);
    assert_eq!(chart.title.as_deref(), Some(" Padded Title "), "a:t whitespace");
    let s = &chart.series[0];
    assert_eq!(s.name.as_deref(), Some(" S!$B$1 "), "c:f whitespace");
    assert_eq!(
        s.values,
        duke_sheets_chart::DataReference::Formula(" S!$B$2:$B$4 ".into()),
        "value c:f whitespace"
    );
    assert_eq!(
        s.data_labels.as_ref().and_then(|d| d.separator.as_deref()),
        Some(" | "),
        "c:separator whitespace"
    );
}
