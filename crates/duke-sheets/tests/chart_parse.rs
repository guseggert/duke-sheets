//! Standard (`c:`) chart parsing: element coverage and form-insensitivity.
//!
//! `<x/>` and `<x></x>` are the same document, and indentation between
//! elements is not content. The parser is a hand-written event loop with
//! a separate arm per event kind, which is exactly where those stop
//! being true, so a fixture covering its vocabulary is run through both
//! rewrites and the results compared.

mod xml_forms;

use duke_sheets_chart::Chart;
use xml_forms::{collapse_indentation, expand_empty_elements};

/// A bar chart using the parser's modelled vocabulary: titles, a legend,
/// both axes with scaling and gridlines, series with formatting, data
/// points, data labels, a trendline, error bars, markers, a data table,
/// 3-D view settings and the chart-space level switches.
const KITCHEN_SINK: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<c:roundedCorners val="0"/>
<c:chart>
<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>My Chart</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:autoTitleDeleted val="0"/>
<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:depthPercent val="100"/><c:rAngAx val="1"/><c:perspective val="30"/></c:view3D>
<c:plotArea>
<c:layout/>
<c:barChart>
<c:barDir val="col"/>
<c:grouping val="clustered"/>
<c:varyColors val="0"/>
<c:ser>
<c:idx val="0"/>
<c:order val="0"/>
<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
<c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
<c:invertIfNegative val="0"/>
<c:dPt><c:idx val="1"/><c:invertIfNegative val="0"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr></c:dPt>
<c:dLbls><c:numFmt formatCode="0.00" sourceLinked="0"/><c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:separator>; </c:separator></c:dLbls>
<c:trendline><c:trendlineType val="linear"/><c:period val="2"/><c:forward val="1"/><c:backward val="1"/><c:intercept val="0"/><c:dispRSqr val="1"/><c:dispEq val="1"/></c:trendline>
<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="0"/><c:val val="2.5"/></c:errBars>
<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser>
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/></c:dLbls>
<c:gapWidth val="219"/>
<c:overlap val="-27"/>
<c:serLines/>
<c:axId val="111111111"/>
<c:axId val="222222222"/>
</c:barChart>
<c:catAx>
<c:axId val="111111111"/>
<c:scaling><c:orientation val="minMax"/></c:scaling>
<c:delete val="0"/>
<c:axPos val="b"/>
<c:majorTickMark val="out"/>
<c:minorTickMark val="none"/>
<c:tickLblPos val="nextTo"/>
<c:crossAx val="222222222"/>
<c:crosses val="autoZero"/>
</c:catAx>
<c:valAx>
<c:axId val="222222222"/>
<c:scaling><c:orientation val="minMax"/><c:max val="10"/><c:min val="0"/></c:scaling>
<c:delete val="0"/>
<c:axPos val="l"/>
<c:majorGridlines/>
<c:minorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:minorGridlines>
<c:numFmt formatCode="General" sourceLinked="1"/>
<c:majorTickMark val="out"/>
<c:minorTickMark val="none"/>
<c:tickLblPos val="nextTo"/>
<c:crossAx val="111111111"/>
<c:crosses val="autoZero"/>
<c:crossBetween val="between"/>
<c:majorUnit val="2"/>
<c:minorUnit val="1"/>
</c:valAx>
<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable>
<c:spPr><a:noFill/></c:spPr>
</c:plotArea>
<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
<c:plotVisOnly val="1"/>
<c:dispBlanksAs val="gap"/>
<c:showDLblsOverMax val="0"/>
</c:chart>
<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></c:spPr>
</c:chartSpace>"#;

fn parse(xml: &str) -> Chart {
    duke_sheets_chart::parse::parse_chart_xml(xml.as_bytes()).expect("parse chart")
}

/// `<x/>` and `<x></x>` are the same document; the parser must not care.
#[test]
fn self_closing_elements_parse_the_same_as_expanded_ones() {
    let expanded = expand_empty_elements(KITCHEN_SINK);
    assert!(
        !expanded.contains("/>"),
        "the rewrite left self-closing tags behind"
    );
    assert_eq!(
        parse(KITCHEN_SINK),
        parse(&expanded),
        "self-closing and expanded forms parsed differently"
    );
}

/// Indentation between elements is not content.
#[test]
fn indentation_between_elements_does_not_change_the_model() {
    let collapsed = collapse_indentation(KITCHEN_SINK);
    assert!(!collapsed.contains(">\n"), "the rewrite left indentation behind");
    assert_eq!(
        parse(KITCHEN_SINK),
        parse(&collapsed),
        "indentation between elements changed the model"
    );
}

/// Characterization: pin what the parser makes of the fixture, so the
/// event loop can be restructured against something concrete.
#[test]
fn kitchen_sink_parses_to_the_expected_model() {
    let chart = parse(KITCHEN_SINK);

    assert_eq!(chart.title.as_deref(), Some("My Chart"));
    assert_eq!(chart.series.len(), 1);

    let s = &chart.series[0];
    assert_eq!(
        s.categories,
        Some(duke_sheets_chart::DataReference::Formula(
            "Sheet1!$A$2:$A$4".into()
        ))
    );
    assert_eq!(
        s.values,
        duke_sheets_chart::DataReference::Formula("Sheet1!$B$2:$B$4".into())
    );
    assert_eq!(s.invert_if_negative, Some(false));
    assert_eq!(
        s.shape_properties
            .as_ref()
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("4472C4")
    );

    let tl = s.trendline.as_ref().expect("trendline");
    assert_eq!(tl.period, Some(2));
    assert_eq!(tl.forward, Some(1.0));
    assert_eq!(tl.backward, Some(1.0));
    assert_eq!(tl.display_r_squared, Some(true));
    assert_eq!(tl.display_equation, Some(true));

    let eb = s.error_bars.as_ref().expect("error bars");
    assert_eq!(eb.value, Some(2.5));
    assert_eq!(eb.no_end_cap, Some(false));

    assert_eq!(s.data_points.len(), 1);
    assert_eq!(s.data_points[0].index, 1);
    assert_eq!(
        s.data_points[0]
            .shape_properties
            .as_ref()
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("ED7D31")
    );

    let dl = s.data_labels.as_ref().expect("series data labels");
    assert_eq!(dl.show_value, Some(true));
    assert_eq!(dl.show_category_name, Some(false));
    assert_eq!(dl.separator.as_deref(), Some("; "));
    assert_eq!(
        dl.number_format.as_ref().map(|nf| nf.format_code.as_str()),
        Some("0.00")
    );

    assert_eq!(chart.auto_title_deleted, Some(false));
    assert_eq!(chart.rounded_corners, Some(false));
    assert_eq!(chart.plot_visible_only, Some(true));
    assert_eq!(chart.show_dlbls_over_max, Some(false));
    assert_eq!(chart.gap_width, Some(219));
    assert_eq!(chart.overlap, Some(-27));
    assert_eq!(chart.vary_colors, Some(false));
    assert!(chart.series_lines.is_some(), "serLines");

    let cat = chart.category_axis.as_ref().expect("category axis");
    assert_eq!(cat.position, Some(duke_sheets_chart::AxisPosition::Bottom));
    assert_eq!(cat.delete, Some(false));
    assert!(!cat.major_gridlines);

    let val = chart.value_axis.as_ref().expect("value axis");
    assert_eq!(val.minimum, Some(0.0));
    assert_eq!(val.maximum, Some(10.0));
    assert_eq!(val.major_unit, Some(2.0));
    assert_eq!(val.minor_unit, Some(1.0));
    assert!(val.major_gridlines, "majorGridlines");
    assert!(val.minor_gridlines, "minorGridlines");
    assert_eq!(
        val.minor_gridlines_shape_properties
            .as_ref()
            .and_then(|sp| sp.line.as_ref())
            .and_then(|l| l.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("D9D9D9")
    );

    let v3 = chart.view_3d.as_ref().expect("view3D");
    assert_eq!(v3.rotate_x, Some(15));
    assert_eq!(v3.rotate_y, Some(20));
    assert_eq!(v3.depth_percent, Some(100));
    assert_eq!(v3.perspective, Some(30));
    assert_eq!(v3.right_angle_axes, Some(true));

    let dt = chart.data_table.as_ref().expect("dTable");
    assert_eq!(dt.show_horizontal_border, Some(true));
    assert_eq!(dt.show_keys, Some(true));

    let legend = chart.legend.as_ref().expect("legend");
    assert_eq!(legend.position, duke_sheets_chart::LegendPosition::Bottom);
    assert!(!legend.overlay);
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
