//! chartEx parsing: element coverage and self-closing equivalence.
//!
//! `<x/>` and `<x></x>` are the same document, so the parser must produce
//! the same model from either. The parser used to have a separate arm for
//! empty elements that had drifted from the one for start tags, which
//! silently dropped whole subtrees - a self-closing `cx:series` left an
//! empty plot area, and a self-closing `cx:layoutPr` or `cx:spPr`
//! vanished. `self_closing_elements_parse_the_same_as_expanded_ones`
//! pins the invariant across the whole element vocabulary.

use duke_sheets_chart::{ChartEx, ChartExLayout, ChartExScaling};

/// Every chartEx element the parser models, bar those whose content is
/// captured as raw bytes (`geoCache`, `valueColors`, `clrMapOvr`,
/// `extLst`, `rich`, `txPr`); expanding a self-closing tag inside those
/// changes the captured bytes, so they are exercised separately.
const KITCHEN_SINK: &str = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<cx:chartData>
<cx:data id="0">
<cx:strDim type="cat"><cx:f>Sheet1!$A$1:$A$3</cx:f><cx:lvl ptCount="3" name="Cats"><cx:pt idx="0">a</cx:pt><cx:pt idx="1">b</cx:pt><cx:pt idx="2">c</cx:pt></cx:lvl></cx:strDim>
<cx:numDim type="val"><cx:f>Sheet1!$B$1:$B$3</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt><cx:pt idx="2">3</cx:pt></cx:lvl></cx:numDim>
</cx:data>
<cx:externalData r:id="rId9" autoUpdate="0"/>
</cx:chartData>
<cx:chart>
<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>My Title</cx:v></cx:txData></cx:tx></cx:title>
<cx:plotArea>
<cx:plotAreaRegion>
<cx:series layoutId="waterfall" uniqueId="{AAAA0000-1111-2222-3333-444455556666}" hidden="0" ownerIdx="1" formatIdx="2">
<cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Series1</cx:v></cx:txData></cx:tx>
<cx:dataId val="0"/>
<cx:dataLabels pos="ctr"><cx:visibility seriesName="0" categoryName="1" value="1"/><cx:numFmt formatCode="0.00" sourceLinked="0"/><cx:separator>;</cx:separator></cx:dataLabels>
<cx:dataPt idx="1"><cx:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></cx:spPr></cx:dataPt>
<cx:layoutPr><cx:visibility connectorLines="1" meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/><cx:statistics quartileMethod="inclusive"/><cx:subtotals><cx:idx val="0"/><cx:idx val="2"/></cx:subtotals></cx:layoutPr>
<cx:axisId>0</cx:axisId>
<cx:axisId>1</cx:axisId>
<cx:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="654321"/></a:solidFill><a:prstDash val="dash"/></a:ln></cx:spPr>
</cx:series>
<cx:series layoutId="clusteredColumn" hidden="1">
<cx:dataId val="0"/>
<cx:layoutPr><cx:parentLabelLayout val="banner"/><cx:regionLabelLayout val="bestFitOnly"/><cx:aggregation/><cx:binning intervalClosed="l" underflow="auto" overflow="auto"><cx:binSize val="2.5"/><cx:binCount val="7"/></cx:binning></cx:layoutPr>
<cx:valueColorPositions count="3"><cx:min><cx:extremeValue/></cx:min><cx:mid><cx:percent val="50"/></cx:mid><cx:max><cx:number val="12.5"/></cx:max></cx:valueColorPositions>
<cx:spPr><a:ln w="19050"><a:noFill/></a:ln></cx:spPr>
</cx:series>
<cx:plotSurface><cx:spPr><a:noFill/></cx:spPr></cx:plotSurface>
</cx:plotAreaRegion>
<cx:axis id="0" hidden="0"><cx:catScaling gapWidth="0.5"/><cx:title><cx:tx><cx:txData><cx:v>Cat Axis</cx:v></cx:txData></cx:tx></cx:title><cx:majorTickMarks type="out"/><cx:minorTickMarks type="none"/><cx:tickLabels/><cx:numFmt formatCode="General" sourceLinked="1"/></cx:axis>
<cx:axis id="1"><cx:valScaling min="0" max="10" majorUnit="2" minorUnit="1"/><cx:units unit="hundreds"/><cx:majorGridlines/><cx:minorGridlines><cx:spPr><a:ln><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></a:ln></cx:spPr></cx:minorGridlines><cx:tickLabels/></cx:axis>
</cx:plotArea>
<cx:legend pos="b" align="ctr" overlay="1"><cx:spPr><a:noFill/></cx:spPr></cx:legend>
</cx:chart>
<cx:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></cx:spPr>
<cx:printSettings>
<cx:headerFooter differentOddEven="0" differentFirst="0"><cx:oddHeader>H</cx:oddHeader><cx:oddFooter>F</cx:oddFooter></cx:headerFooter>
<cx:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
<cx:pageSetup paperSize="9" orientation="landscape" blackAndWhite="0" draft="0" useFirstPageNumber="1" firstPageNumber="3" horizontalDpi="600" verticalDpi="600" copies="2"/>
</cx:printSettings>
</cx:chartSpace>"#;

fn parse(xml: &str) -> ChartEx {
    duke_sheets_chart::parse::parse_chart_ex_xml(xml.as_bytes()).expect("parse chartEx")
}

/// Rewrite every self-closing tag `<x .../>` as `<x ...></x>`, respecting
/// quoted attribute values. `<?xml?>` and comments are left alone.
fn expand_empty_elements(xml: &str) -> String {
    let b = xml.as_bytes();
    let mut out = String::with_capacity(xml.len() + 256);
    let mut i = 0usize;
    while i < b.len() {
        if b[i] != b'<' || i + 1 >= b.len() || matches!(b[i + 1], b'?' | b'!' | b'/') {
            out.push(b[i] as char);
            i += 1;
            continue;
        }
        // Scan to the matching '>', skipping quoted regions.
        let start = i;
        let mut j = i + 1;
        let mut quote: Option<u8> = None;
        while j < b.len() {
            match (quote, b[j]) {
                (Some(q), c) if c == q => quote = None,
                (None, c @ (b'"' | b'\'')) => quote = Some(c),
                (None, b'>') => break,
                _ => {}
            }
            j += 1;
        }
        if j >= b.len() {
            out.push_str(&xml[start..]);
            break;
        }
        let tag = &xml[start..=j];
        if tag.ends_with("/>") {
            let inner = &tag[1..tag.len() - 2];
            let name_end = inner
                .find(|c: char| c.is_whitespace())
                .unwrap_or(inner.len());
            let name = &inner[..name_end];
            out.push('<');
            out.push_str(inner.trim_end());
            out.push('>');
            out.push_str("</");
            out.push_str(name);
            out.push('>');
        } else {
            out.push_str(tag);
        }
        i = j + 1;
    }
    out
}

/// `<x/>` and `<x></x>` are the same document; the parser must not care.
#[test]
fn self_closing_elements_parse_the_same_as_expanded_ones() {
    let expanded = expand_empty_elements(KITCHEN_SINK);
    assert!(
        !expanded.contains("/>"),
        "expander left self-closing tags behind"
    );
    assert_eq!(
        parse(KITCHEN_SINK),
        parse(&expanded),
        "self-closing and expanded forms parsed differently"
    );
}

/// The inverse direction: a document authored entirely with expanded
/// tags must survive being collapsed where the writer would self-close.
#[test]
fn self_closing_subtrees_are_not_dropped() {
    // Each of these used to be dropped in its self-closing form because
    // the parser only acted on the element's End event.
    let cases: &[(&str, &str)] = &[
        ("series", r#"<cx:series layoutId="waterfall"/>"#),
        ("layoutPr", r#"<cx:layoutPr/>"#),
        ("spPr", r#"<cx:spPr/>"#),
        ("dataLabels", r#"<cx:dataLabels pos="ctr"/>"#),
    ];
    for (name, snippet) in cases {
        let doc = match *name {
            "series" => format!(
                r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart><cx:plotArea><cx:plotAreaRegion>{snippet}</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
            ),
            _ => format!(
                r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:dataId val="0"/>{snippet}</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
            ),
        };
        let expanded = expand_empty_elements(&doc);
        assert_eq!(
            parse(&doc),
            parse(&expanded),
            "self-closing cx:{name} parsed differently from its expanded form"
        );
        assert_eq!(
            parse(&doc).plot_area.series.len(),
            1,
            "self-closing cx:{name} dropped the series"
        );
    }
}

/// Characterization: pin what the parser makes of a document covering
/// the modelled element vocabulary, so a refactor of the event loop has
/// to preserve it.
#[test]
fn kitchen_sink_parses_to_the_expected_model() {
    let cx = parse(KITCHEN_SINK);

    // chartData
    assert_eq!(cx.data.len(), 1);
    let data = &cx.data[0];
    assert_eq!(data.id, 0);
    assert_eq!(data.dimensions.len(), 2);
    match &data.dimensions[0] {
        duke_sheets_chart::ChartExDimension::String {
            formula, levels, ..
        } => {
            assert_eq!(formula.as_deref(), Some("Sheet1!$A$1:$A$3"));
            assert_eq!(levels[0].points.len(), 3);
            assert_eq!(levels[0].name.as_deref(), Some("Cats"));
        }
        other => panic!("first dimension should be a strDim, got {other:?}"),
    }
    match &data.dimensions[1] {
        duke_sheets_chart::ChartExDimension::Numeric {
            formula, levels, ..
        } => {
            assert_eq!(formula.as_deref(), Some("Sheet1!$B$1:$B$3"));
            assert_eq!(levels[0].format_code.as_deref(), Some("General"));
        }
        other => panic!("second dimension should be a numDim, got {other:?}"),
    }
    let ext = cx.external_data.as_ref().expect("externalData");
    assert_eq!(ext.rel_id, "rId9");
    assert_eq!(ext.auto_update, Some(false));

    // title and legend
    assert_eq!(
        cx.title.as_ref().and_then(|t| t.text.clone()).as_deref(),
        Some("My Title")
    );
    let legend = cx.legend.as_ref().expect("legend");
    assert_eq!(legend.position.as_deref(), Some("b"));
    assert_eq!(legend.overlay, Some(true));

    // series
    let s = &cx.plot_area.series[0];
    assert_eq!(s.layout, ChartExLayout::Waterfall);
    assert_eq!(s.data_id, 0);
    assert_eq!(s.axis_ids, vec![0, 1]);
    assert_eq!(s.data_points.len(), 1);
    assert_eq!(s.data_points[0].idx, 1);
    let labels = s.data_labels.as_ref().expect("dataLabels");
    assert_eq!(labels.separator.as_deref(), Some(";"));
    let lp = s.layout_properties.as_ref().expect("layoutPr");
    assert_eq!(lp.subtotals, Some(vec![0, 2]));
    assert_eq!(
        lp.statistics.as_ref().and_then(|s| s.quartile_method.clone()),
        Some("inclusive".to_string())
    );
    let vis = lp.visibility.as_ref().expect("visibility");
    assert_eq!(vis.connector_lines, Some(true));
    assert_eq!(vis.mean_line, Some(false));

    // axes
    assert_eq!(cx.plot_area.axes.len(), 2);
    let cat_axis = &cx.plot_area.axes[0];
    assert!(matches!(
        cat_axis.scaling,
        ChartExScaling::Category {
            gap_width: Some(_)
        }
    ));
    assert_eq!(cat_axis.major_tick_marks.as_deref(), Some("out"));
    assert!(cat_axis.tick_labels);
    let val_axis = &cx.plot_area.axes[1];
    assert!(matches!(val_axis.scaling, ChartExScaling::Value { .. }));
    // A bare cx:majorGridlines carries no override; the minor one does.
    let major = val_axis.major_gridlines.as_ref().expect("majorGridlines");
    assert!(major.shape_properties.is_none());
    let minor = val_axis.minor_gridlines.as_ref().expect("minorGridlines");
    assert!(minor.shape_properties.is_some());

    // print settings
    let ps = cx.print_settings.as_ref().expect("printSettings");
    let pm = ps.page_margins.as_ref().expect("pageMargins");
    assert_eq!(pm.left, Some(0.7));
    let setup = ps.page_setup.as_ref().expect("pageSetup");
    assert_eq!(setup.orientation.as_deref(), Some("landscape"));
    assert_eq!(setup.copies, Some(2));
    assert_eq!(
        ps.header_footer.as_ref().and_then(|h| h.odd_header.clone()),
        Some("H".to_string())
    );
}

/// The kitchen sink must survive a write/read cycle unchanged, which is
/// what makes it usable as a refactor guard for the writer too.
#[test]
fn kitchen_sink_survives_a_round_trip() {
    let first = parse(KITCHEN_SINK);
    let bytes = duke_sheets_chart::write::chart_ex_part_bytes(&first).expect("write");
    let second =
        duke_sheets_chart::parse::parse_chart_ex_xml(&bytes[..]).expect("reparse written chartEx");
    assert_eq!(first, second, "chartEx round trip is not idempotent");
}

/// Raw-capture regions are the one place a self-closing element is not
/// split: they replay source bytes, so `<x/>` must stay `<x/>` rather
/// than become `<x></x>`.
#[test]
fn raw_captured_regions_keep_self_closing_tags_verbatim() {
    let doc = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:dataId val="0"/><cx:layoutPr><cx:geography cultureLanguage="en-US"><cx:geoCache provider="x"><cx:binary/><cx:leaf a="1"/></cx:geoCache></cx:geography></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;

    let cx = parse(doc);
    let cache = cx.plot_area.series[0]
        .layout_properties
        .as_ref()
        .and_then(|l| l.geography.as_ref())
        .and_then(|g| g.raw_geo_cache.as_ref())
        .expect("geoCache captured");
    let text = String::from_utf8_lossy(cache);
    assert!(
        text.contains("<cx:binary/>") && text.contains(r#"<cx:leaf a="1"/>"#),
        "self-closing tags inside a raw capture were rewritten: {text}"
    );

    // And the capture survives a write/read cycle unchanged.
    let bytes = duke_sheets_chart::write::chart_ex_part_bytes(&cx).expect("write");
    let again = duke_sheets_chart::parse::parse_chart_ex_xml(&bytes[..]).expect("reparse");
    assert_eq!(cx, again, "raw geoCache capture changed on round trip");
}
