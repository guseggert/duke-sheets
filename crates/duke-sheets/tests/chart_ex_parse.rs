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

/// Every chartEx element the parser models, plus the subtrees it skips
/// (`txPr`, `clrMapOvr`, `fmtOvrs`, `extLst`), whose presence must not
/// disturb what surrounds them. Only `geoCache` and `valueColors`
/// capture raw bytes; expanding a self-closing tag inside those changes
/// the captured bytes, so they are exercised separately.
const KITCHEN_SINK: &str = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<cx:chartData>
<cx:data id="0">
<cx:strDim type="cat"><cx:f>Sheet1!$A$1:$A$3</cx:f><cx:lvl ptCount="3" name="Cats"><cx:pt idx="0">a</cx:pt><cx:pt idx="1">b</cx:pt><cx:pt idx="2">c</cx:pt></cx:lvl></cx:strDim>
<cx:numDim type="val"><cx:f>Sheet1!$B$1:$B$3</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt><cx:pt idx="2">3</cx:pt></cx:lvl></cx:numDim>
</cx:data>
<cx:externalData r:id="rId9" autoUpdate="0"/>
</cx:chartData>
<cx:chart>
<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>My Title</cx:v></cx:txData></cx:tx><cx:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></cx:spPr><cx:offset t="0.11" l="0.22"/></cx:title>
<cx:plotArea>
<cx:plotAreaRegion>
<cx:series layoutId="waterfall" uniqueId="{AAAA0000-1111-2222-3333-444455556666}" hidden="0" ownerIdx="1" formatIdx="2">
<cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Series1</cx:v></cx:txData></cx:tx>
<cx:dataId val="0"/>
<cx:dataLabels pos="ctr"><cx:numFmt formatCode="0.00" sourceLinked="0"/><cx:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></cx:spPr><cx:visibility seriesName="0" categoryName="1" value="1"/><cx:separator>;</cx:separator><cx:dataLabel idx="1" pos="r"><cx:numFmt formatCode="0.000" sourceLinked="0"/><cx:spPr><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></cx:spPr><cx:visibility seriesName="1" categoryName="0" value="1"/><cx:separator>|</cx:separator></cx:dataLabel><cx:dataLabelHidden idx="2"/></cx:dataLabels>
<cx:dataPt idx="1"><cx:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></cx:spPr></cx:dataPt>
<cx:layoutPr><cx:visibility connectorLines="1" meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/><cx:statistics quartileMethod="inclusive"/><cx:subtotals><cx:idx val="0"/><cx:idx val="2"/></cx:subtotals></cx:layoutPr>
<cx:axisId>0</cx:axisId>
<cx:axisId>1</cx:axisId>
<cx:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="654321"/></a:solidFill><a:prstDash val="dash"/></a:ln><a:extLst><a:ext uri="{FF00}"><a:dummy/></a:ext></a:extLst></cx:spPr>
</cx:series>
<cx:series layoutId="clusteredColumn" hidden="1">
<cx:dataId val="0"/>
<cx:layoutPr><cx:parentLabelLayout val="banner"/><cx:regionLabelLayout val="bestFitOnly"/><cx:aggregation/><cx:binning intervalClosed="l" underflow="auto" overflow="auto"><cx:binSize val="2.5"/><cx:binCount val="7"/></cx:binning></cx:layoutPr>
<cx:valueColorPositions count="3"><cx:min><cx:extremeValue/></cx:min><cx:mid><cx:percent val="50"/></cx:mid><cx:max><cx:number val="12.5"/></cx:max></cx:valueColorPositions>
<cx:spPr><a:ln w="19050"><a:noFill/></a:ln></cx:spPr>
</cx:series>
<cx:plotSurface><cx:spPr><a:noFill/></cx:spPr></cx:plotSurface>
</cx:plotAreaRegion>
<cx:axis id="0" hidden="0"><cx:catScaling gapWidth="0.5"/><cx:title><cx:tx><cx:txData><cx:v>Cat Axis</cx:v></cx:txData></cx:tx><cx:spPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></cx:spPr><cx:offset t="0.33" l="0.44"/></cx:title><cx:majorTickMarks type="out"/><cx:minorTickMarks type="none"/><cx:tickLabels/><cx:numFmt formatCode="General" sourceLinked="1"/></cx:axis>
<cx:axis id="1"><cx:valScaling min="0" max="10" majorUnit="2" minorUnit="1"/><cx:units unit="hundreds"/><cx:majorGridlines/><cx:minorGridlines><cx:spPr><a:ln><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></a:ln><a:extLst><a:ext uri="{GG00}"/></a:extLst></cx:spPr></cx:minorGridlines><cx:tickLabels/></cx:axis>
</cx:plotArea>
<cx:legend pos="b" align="ctr" overlay="1"><cx:spPr><a:noFill/></cx:spPr><cx:offset t="0.55" l="0.66"/></cx:legend>
</cx:chart>
<cx:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></cx:spPr>
<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>t</a:t></a:r></a:p></cx:txPr>
<cx:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<cx:fmtOvrs><cx:fmtOvr idx="0"><cx:spPr><a:noFill/></cx:spPr></cx:fmtOvr></cx:fmtOvrs>
<cx:printSettings>
<cx:headerFooter differentOddEven="0" differentFirst="0"><cx:oddHeader>H</cx:oddHeader><cx:oddFooter>F</cx:oddFooter></cx:headerFooter>
<cx:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
<cx:pageSetup paperSize="9" orientation="landscape" blackAndWhite="0" draft="0" useFirstPageNumber="1" firstPageNumber="3" horizontalDpi="600" verticalDpi="600" copies="2"/>
</cx:printSettings>
<cx:extLst><cx:ext uri="{AA00}"><cx:leaf/></cx:ext></cx:extLst>
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
    assert_eq!(
        cx.title
            .as_ref()
            .and_then(|t| t.shape_properties.as_ref())
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("112233"),
        "chart title spPr must land on the title"
    );
    assert_eq!(
        cx.title.as_ref().and_then(|t| t.offset.clone()),
        Some(duke_sheets_chart::ChartExOffset {
            top: Some(0.11),
            left: Some(0.22)
        })
    );
    let legend = cx.legend.as_ref().expect("legend");
    assert_eq!(legend.position.as_deref(), Some("b"));
    assert_eq!(
        legend.offset,
        Some(duke_sheets_chart::ChartExOffset {
            top: Some(0.55),
            left: Some(0.66)
        })
    );
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
    assert_eq!(
        labels
            .shape_properties
            .as_ref()
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("0000FF"),
        "dataLabels spPr must be modelled, not parsed and dropped"
    );
    assert_eq!(
        labels.number_format.as_ref().map(|nf| nf.format_code.as_str()),
        Some("0.00"),
        "a per-label override must not replace the series-level numFmt"
    );
    assert_eq!(labels.hidden_labels, vec![2]);
    assert_eq!(labels.overrides.len(), 1, "cx:dataLabel override must be modelled");
    let ovr = &labels.overrides[0];
    assert_eq!(ovr.idx, 1);
    assert_eq!(ovr.position.as_deref(), Some("r"));
    assert_eq!(
        ovr.number_format.as_ref().map(|nf| nf.format_code.as_str()),
        Some("0.000")
    );
    assert_eq!(
        ovr.shape_properties
            .as_ref()
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("FF00FF")
    );
    assert_eq!(ovr.separator.as_deref(), Some("|"));
    assert_eq!(ovr.visibility_series_name, Some(true));
    assert_eq!(ovr.visibility_category_name, Some(false));
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
    assert_eq!(
        cat_axis
            .title
            .as_ref()
            .and_then(|t| t.shape_properties.as_ref())
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("445566"),
        "axis title spPr must land on the title"
    );
    assert_eq!(
        cat_axis.title.as_ref().and_then(|t| t.offset.clone()),
        Some(duke_sheets_chart::ChartExOffset {
            top: Some(0.33),
            left: Some(0.44)
        })
    );
    assert!(
        cat_axis.shape_properties.is_none(),
        "axis title spPr must not be attributed to the axis itself"
    );
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

    // format overrides
    assert_eq!(cx.format_overrides.len(), 1, "cx:fmtOvr must be modelled");
    assert_eq!(cx.format_overrides[0].idx, 0);
    assert!(
        cx.format_overrides[0]
            .shape_properties
            .as_ref()
            .is_some_and(|sp| sp.no_fill),
        "fmtOvr spPr must be read"
    );

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

/// Same contract for the other capture region: `minColor`/`midColor`/
/// `maxColor` bodies inside `cx:valueColors` are raw bytes.
#[test]
fn value_color_captures_keep_self_closing_tags_verbatim() {
    let doc = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:dataId val="0"/><cx:valueColors><cx:minColor><a:srgbClr val="FF0000"/></cx:minColor><cx:maxColor><a:schemeClr val="accent1"><a:lumMod val="60000"/></a:schemeClr></cx:maxColor></cx:valueColors></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;

    let cx = parse(doc);
    let vc = cx.plot_area.series[0]
        .value_colors
        .as_ref()
        .expect("valueColors captured");
    let min = String::from_utf8_lossy(vc.min_color.as_ref().expect("minColor"));
    assert!(
        min.contains(r#"<a:srgbClr val="FF0000"/>"#),
        "self-closing tag inside minColor was rewritten: {min}"
    );
    assert!(vc.mid_color.is_none());
    let max = String::from_utf8_lossy(vc.max_color.as_ref().expect("maxColor"));
    assert!(max.contains(r#"<a:lumMod val="60000"/>"#), "maxColor: {max}");

    let bytes = duke_sheets_chart::write::chart_ex_part_bytes(&cx).expect("write");
    let again = duke_sheets_chart::parse::parse_chart_ex_xml(&bytes[..]).expect("reparse");
    assert_eq!(cx, again, "valueColors capture changed on round trip");
}

/// A self-closing capture opener is split before its capture writer
/// exists, so the captured bytes come out expanded. That is a semantic
/// no-op; this pins it as intended behaviour rather than an accident.
#[test]
fn self_closing_capture_opener_round_trips_as_expanded_form() {
    let doc = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:dataId val="0"/><cx:layoutPr><cx:geography cultureLanguage="en-US"><cx:geoCache provider="x"/></cx:geography></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;

    let cx = parse(doc);
    let cache = cx.plot_area.series[0]
        .layout_properties
        .as_ref()
        .and_then(|l| l.geography.as_ref())
        .and_then(|g| g.raw_geo_cache.as_ref())
        .expect("empty geoCache still captured");
    assert_eq!(
        String::from_utf8_lossy(cache),
        r#"<cx:geoCache provider="x"></cx:geoCache>"#
    );

    let bytes = duke_sheets_chart::write::chart_ex_part_bytes(&cx).expect("write");
    let again = duke_sheets_chart::parse::parse_chart_ex_xml(&bytes[..]).expect("reparse");
    assert_eq!(cx, again, "empty geoCache changed on round trip");
}

/// `version`, `featureList` and `fallbackImg` are `CT_ChartSpace`
/// attributes; both sides used to drop them, so they vanished from any
/// file that carried them.
#[test]
fn chart_space_attributes_round_trip() {
    let doc = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" version="1.5" featureList="waterfall" fallbackImg="rId7"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;

    let cx = parse(doc);
    assert_eq!(cx.version.as_deref(), Some("1.5"));
    assert_eq!(cx.feature_list.as_deref(), Some("waterfall"));
    assert_eq!(cx.fallback_img.as_deref(), Some("rId7"));

    let bytes = duke_sheets_chart::write::chart_ex_part_bytes(&cx).expect("write");
    let again = duke_sheets_chart::parse::parse_chart_ex_xml(&bytes[..]).expect("reparse");
    assert_eq!(cx, again, "chartSpace attributes lost on round trip");
}

/// `a:extLst` is a valid last child of `a:CT_ShapeProperties`. Entering
/// its skip region must not desynchronize the `cx:spPr` nesting depth:
/// when it did, the containing spPr never closed, its formatting was
/// lost, and every later spPr in the document was dropped too.
#[test]
fn ext_lst_inside_shape_properties_does_not_poison_later_ones() {
    for ext in ["<a:extLst><a:ext uri=\"{X}\"/></a:extLst>", "<a:extLst/>"] {
        let doc = format!(
            r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:dataId val="0"/><cx:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>{ext}</cx:spPr></cx:series><cx:series layoutId="waterfall"><cx:dataId val="0"/><cx:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></cx:spPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
        );
        let cx = parse(&doc);
        let fill = |i: usize| {
            cx.plot_area.series[i]
                .shape_properties
                .as_ref()
                .and_then(|sp| sp.solid_fill.as_ref())
                .map(|c| c.hex.as_str().to_owned())
        };
        assert_eq!(fill(0), Some("FF0000".into()), "first spPr lost ({ext})");
        assert_eq!(fill(1), Some("00FF00".into()), "later spPr lost ({ext})");
    }
}

/// The writer must emit CT_Series children in schema order; the reader
/// is order-blind, so round-trip equality cannot see a violation.
#[test]
fn series_children_are_written_in_schema_order() {
    let doc = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:tx><cx:txData><cx:v>S</cx:v></cx:txData></cx:tx><cx:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></cx:spPr><cx:valueColors><cx:minColor><a:srgbClr val="FF0000"/></cx:minColor></cx:valueColors><cx:valueColorPositions count="2"><cx:min><cx:extremeValue/></cx:min></cx:valueColorPositions><cx:dataPt idx="0"><cx:spPr><a:noFill/></cx:spPr></cx:dataPt><cx:dataLabels pos="ctr"><cx:numFmt formatCode="0.0" sourceLinked="0"/><cx:spPr><a:noFill/></cx:spPr><cx:visibility value="1"/><cx:separator>;</cx:separator><cx:dataLabel idx="0" pos="r"><cx:numFmt formatCode="0.00" sourceLinked="0"/><cx:spPr><a:noFill/></cx:spPr><cx:visibility value="0"/><cx:separator>,</cx:separator></cx:dataLabel></cx:dataLabels><cx:dataId val="0"/><cx:layoutPr><cx:subtotals/></cx:layoutPr><cx:axisId>0</cx:axisId></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;

    let cx = parse(doc);
    let out = String::from_utf8(duke_sheets_chart::write::chart_ex_part_bytes(&cx).unwrap()).unwrap();

    let series = &out[out.find("<cx:series").unwrap()..out.find("</cx:series>").unwrap()];
    let pos = |needle: &str| {
        series
            .find(needle)
            .unwrap_or_else(|| panic!("{needle} missing from written series: {series}"))
    };
    // CT_Series sequence.
    let order = [
        "<cx:tx>",
        "<cx:spPr>",
        "<cx:valueColors>",
        "<cx:valueColorPositions",
        "<cx:dataPt",
        "<cx:dataLabels",
        "<cx:dataId",
        "<cx:layoutPr>",
        "<cx:axisId>",
    ];
    for pair in order.windows(2) {
        assert!(
            pos(pair[0]) < pos(pair[1]),
            "{} must precede {} in CT_Series; got: {series}",
            pair[0],
            pair[1]
        );
    }

    // CT_DataLabel sequence inside the override.
    let lbl = &series[pos("<cx:dataLabel ")..series.find("</cx:dataLabel>").unwrap()];
    assert!(
        lbl.contains(r#"idx="0""#) && lbl.contains(r#"pos="r""#),
        "override must keep idx and pos attributes: {lbl}"
    );
    let lpos = |needle: &str| lbl.find(needle).unwrap_or_else(|| panic!("{needle} missing: {lbl}"));
    for pair in ["<cx:numFmt", "<cx:spPr>", "<cx:visibility", "<cx:separator>"].windows(2) {
        assert!(
            lpos(pair[0]) < lpos(pair[1]),
            "{} must precede {} in CT_DataLabel; got: {lbl}",
            pair[0],
            pair[1]
        );
    }

    // CT_ChartTitle sequence: tx then spPr.
    let cx2 = parse(KITCHEN_SINK);
    let out2 = String::from_utf8(duke_sheets_chart::write::chart_ex_part_bytes(&cx2).unwrap()).unwrap();
    let title = &out2[out2.find("<cx:title").unwrap()..out2.find("</cx:title>").unwrap()];
    assert!(
        title.find("<cx:tx>").unwrap() < title.find("<cx:spPr>").unwrap(),
        "tx must precede spPr in CT_ChartTitle: {title}"
    );
}

/// A captured subtree is replayed verbatim on write, so every event kind
/// inside it has to be kept - comments and CDATA included, which the
/// capture path used to drop while the skip path never had to care.
#[test]
fn raw_captures_keep_comments_and_cdata() {
    let doc = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:dataId val="0"/><cx:layoutPr><cx:geography cultureLanguage="en-US"><cx:geoCache provider="x"><!-- keep me --><![CDATA[raw&bytes]]><cx:leaf a="1"/></cx:geoCache></cx:geography></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;

    let cx = parse(doc);
    let cache = cx.plot_area.series[0]
        .layout_properties
        .as_ref()
        .and_then(|l| l.geography.as_ref())
        .and_then(|g| g.raw_geo_cache.as_ref())
        .expect("geoCache captured");
    let text = String::from_utf8_lossy(cache);
    assert!(text.contains("<!-- keep me -->"), "comment dropped: {text}");
    assert!(text.contains("<![CDATA[raw&bytes]]>"), "CDATA dropped: {text}");
    assert!(text.contains(r#"<cx:leaf a="1"/>"#), "element dropped: {text}");

    let bytes = duke_sheets_chart::write::chart_ex_part_bytes(&cx).expect("write");
    let again = duke_sheets_chart::parse::parse_chart_ex_xml(&bytes[..]).expect("reparse");
    assert_eq!(cx, again, "capture changed on round trip");
}

/// An opaque subtree must not leak nesting depth into the `cx:spPr` it
/// sits inside. The skip half of that is reachable from schema-valid
/// input (`a:extLst`); the capture half is not, but both share one
/// mechanism now and the invariant is what keeps them honest.
#[test]
fn a_capture_inside_shape_properties_does_not_leak_depth() {
    let doc = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:dataId val="0"/><cx:layoutPr><cx:geography cultureLanguage="en-US"><cx:spPr><cx:geoCache provider="x"><cx:leaf/></cx:geoCache><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></cx:spPr></cx:geography></cx:layoutPr></cx:series><cx:series layoutId="waterfall"><cx:dataId val="0"/><cx:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></cx:spPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;

    let cx = parse(doc);
    assert_eq!(
        cx.plot_area.series[1]
            .shape_properties
            .as_ref()
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("00FF00"),
        "a later spPr was lost, so the capture leaked depth"
    );
}

/// `cx:rich` and `cx:clrMapOvr` are kept as bytes for replay. The writer
/// could already emit both; the parser dropped them, so rich titles and
/// colour-map overrides vanished from every file read.
#[test]
fn rich_text_and_color_map_override_round_trip() {
    let doc = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:rich><a:bodyPr rot="60000"/><a:p><a:r><a:t>Rich Title</a:t></a:r></a:p></cx:rich></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:tx><cx:rich><a:p><a:r><a:t>Rich Series</a:t></a:r></a:p></cx:rich></cx:tx><cx:dataId val="0"/></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:title><cx:tx><cx:rich><a:p><a:r><a:t>Rich Axis</a:t></a:r></a:p></cx:rich></cx:tx></cx:title></cx:axis></cx:plotArea></cx:chart><cx:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></cx:chartSpace>"#;

    let cx = parse(doc);

    let title_rich = cx.title.as_ref().and_then(|t| t.rich_text.as_ref()).expect("title rich");
    let title_rich = String::from_utf8_lossy(title_rich);
    assert!(title_rich.starts_with("<cx:rich>"), "rich must be the whole element: {title_rich}");
    assert!(title_rich.contains("Rich Title") && title_rich.contains(r#"rot="60000""#));

    let series_rich = cx.plot_area.series[0]
        .text
        .as_ref()
        .and_then(|t| t.rich.as_ref())
        .expect("series rich");
    assert!(String::from_utf8_lossy(series_rich).contains("Rich Series"));

    let axis_rich = cx.plot_area.axes[0]
        .title
        .as_ref()
        .and_then(|t| t.text.as_ref())
        .and_then(|t| t.rich.as_ref())
        .expect("axis title rich");
    assert!(String::from_utf8_lossy(axis_rich).contains("Rich Axis"));

    // The colour map is carried entirely by attributes, so the capture
    // has to include the element itself, not just its (empty) content.
    let cmo = cx.color_map_override.as_ref().expect("clrMapOvr");
    let cmo = String::from_utf8_lossy(cmo);
    assert!(cmo.contains(r#"bg1="lt1""#) && cmo.contains(r#"folHlink="folHlink""#), "{cmo}");

    let bytes = duke_sheets_chart::write::chart_ex_part_bytes(&cx).expect("write");
    let again = duke_sheets_chart::parse::parse_chart_ex_xml(&bytes[..]).expect("reparse");
    assert_eq!(cx, again, "rich / clrMapOvr changed on round trip");
}
