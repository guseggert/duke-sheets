//! Default chart style and colour style parts for chartEx charts.
//!
//! Excel will not open a workbook whose chartEx part lacks a sibling
//! chart style part and chart colour style part, and it validates the
//! style part against `CT_ChartStyle` ([MS-ODRAWXML] §5.15): every
//! required `CT_StyleEntry` must be present, in schema order, each with
//! its own required `lnRef`, `fillRef`, `effectRef` and `fontRef`
//! children. A workbook missing either part, or carrying a style part
//! with entries missing or reordered, is rejected outright rather than
//! repaired.
//!
//! One divergence from the published schema, established by driving
//! Excel: `id` is `use="optional"` in `CT_ChartStyle`, but Excel refuses
//! a style part that omits it. It is always emitted here.
//!
//! Charts read from a file keep their original parts verbatim; these
//! defaults are for charts built through the model, which have none.
//!
//! These are fixed templates with a single substituted value, so they
//! are built as strings rather than driven through the XML writer used
//! for the chartEx part itself; that keeps them auditable line by line
//! against what Excel emits.

const NS: &str = concat!(
    r#"xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#
);

const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// Hairline in the "15% tint of text 1" that Excel uses for chart
/// borders, gridlines and axis lines.
const SUBTLE_LINE: &str = concat!(
    r#"<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>"#,
    r#"<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>"#,
    r#"</a:schemeClr></a:solidFill><a:round/></a:ln>"#
);

/// Text 1 at 65% luminance, Excel's colour for secondary chart text.
const MUTED_TEXT: &str = concat!(
    r#"<a:schemeClr val="tx1"><a:lumMod val="65000"/>"#,
    r#"<a:lumOff val="35000"/></a:schemeClr>"#
);

const PLAIN_TEXT: &str = r#"<a:schemeClr val="tx1"/>"#;

/// A `CT_StyleEntry` in the generated default.
struct Entry {
    /// Element name, which is also its position in the `CT_ChartStyle` sequence.
    name: &'static str,
    /// `mods` attribute, when Excel sets one.
    mods: Option<&'static str>,
    /// Inner XML of `cs:fillRef`, empty for the bare `idx="0"` form.
    fill_ref_inner: &'static str,
    /// Inner XML of `cs:fontRef`.
    font_ref_inner: &'static str,
    /// Inner XML of `cs:spPr`, when the entry carries a formatting override.
    sp_pr_inner: Option<&'static str>,
    /// `sz` on `cs:defRPr`, in hundredths of a point.
    def_rpr_sz: Option<u32>,
}

impl Entry {
    const fn plain(name: &'static str) -> Self {
        Self {
            name,
            mods: None,
            fill_ref_inner: "",
            font_ref_inner: PLAIN_TEXT,
            sp_pr_inner: None,
            def_rpr_sz: None,
        }
    }

    /// An entry whose only job is to carry text defaults.
    const fn text(name: &'static str, sz: u32) -> Self {
        Self {
            font_ref_inner: MUTED_TEXT,
            def_rpr_sz: Some(sz),
            ..Self::plain(name)
        }
    }

    /// An entry drawn with the subtle hairline (gridlines, leader lines).
    const fn hairline(name: &'static str) -> Self {
        Self {
            sp_pr_inner: Some(SUBTLE_LINE),
            ..Self::plain(name)
        }
    }
}

/// The `CT_ChartStyle` sequence. Order is load-bearing: Excel rejects a
/// style part whose entries are reordered, and rejects one with any
/// required entry missing.
const ENTRIES: &[Entry] = &[
    Entry::text("axisTitle", 900),
    Entry {
        sp_pr_inner: Some(SUBTLE_LINE),
        ..Entry::text("categoryAxis", 900)
    },
    Entry {
        mods: Some("allowNoFillOverride allowNoLineOverride"),
        sp_pr_inner: Some(concat!(
            r#"<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>"#,
            // SUBTLE_LINE, inlined: concat! needs literals.
            r#"<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>"#,
            r#"<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>"#,
            r#"</a:schemeClr></a:solidFill><a:round/></a:ln>"#
        )),
        def_rpr_sz: Some(1000),
        ..Entry::plain("chartArea")
    },
    Entry::text("dataLabel", 900),
    Entry {
        sp_pr_inner: Some(SUBTLE_LINE),
        ..Entry::text("dataLabelCallout", 900)
    },
    // Series colours come from the colour style part: `styleClr val="auto"`
    // takes the next colour in the cycle and `phClr` resolves to it.
    Entry {
        fill_ref_inner: r#"<cs:styleClr val="auto"/>"#,
        sp_pr_inner: Some(r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#),
        ..Entry::plain("dataPoint")
    },
    Entry {
        fill_ref_inner: r#"<cs:styleClr val="auto"/>"#,
        sp_pr_inner: Some(r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#),
        ..Entry::plain("dataPoint3D")
    },
    Entry {
        fill_ref_inner: r#"<cs:styleClr val="auto"/>"#,
        sp_pr_inner: Some(concat!(
            r#"<a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/>"#,
            r#"</a:solidFill><a:round/></a:ln>"#
        )),
        ..Entry::plain("dataPointLine")
    },
    Entry {
        fill_ref_inner: r#"<cs:styleClr val="auto"/>"#,
        sp_pr_inner: Some(concat!(
            r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#,
            r#"<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>"#
        )),
        ..Entry::plain("dataPointMarker")
    },
    // dataPointMarkerLayout is CT_MarkerLayout, not CT_StyleEntry; it is
    // emitted separately, in this position.
    Entry {
        fill_ref_inner: r#"<cs:styleClr val="auto"/>"#,
        sp_pr_inner: Some(concat!(
            r#"<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>"#,
            r#"</a:solidFill></a:ln>"#
        )),
        ..Entry::plain("dataPointWireframe")
    },
    Entry {
        sp_pr_inner: Some(SUBTLE_LINE),
        ..Entry::text("dataTable", 900)
    },
    Entry::plain("downBar"),
    Entry::hairline("dropLine"),
    Entry::hairline("errorBar"),
    Entry::plain("floor"),
    Entry::hairline("gridlineMajor"),
    Entry::hairline("gridlineMinor"),
    Entry::hairline("hiLoLine"),
    Entry::hairline("leaderLine"),
    Entry::text("legend", 900),
    Entry {
        mods: Some("allowNoFillOverride allowNoLineOverride"),
        ..Entry::plain("plotArea")
    },
    Entry {
        mods: Some("allowNoFillOverride allowNoLineOverride"),
        ..Entry::plain("plotArea3D")
    },
    Entry {
        sp_pr_inner: Some(SUBTLE_LINE),
        ..Entry::text("seriesAxis", 900)
    },
    Entry::hairline("seriesLine"),
    Entry::text("title", 1400),
    Entry::hairline("trendline"),
    Entry::text("trendlineLabel", 900),
    Entry::plain("upBar"),
    Entry::text("valueAxis", 900),
    Entry::plain("wall"),
];

/// Marker layout, emitted between `dataPointMarker` and `dataPointWireframe`.
const MARKER_LAYOUT: &str = r#"<cs:dataPointMarkerLayout symbol="circle" size="5"/>"#;

/// Style id emitted on the generated `cs:chartStyle`.
///
/// The id selects which entry of Excel's chart-style gallery is shown as
/// current; it does not affect how the chart renders, since the entries
/// below carry the formatting. Excel accepts any value but requires the
/// attribute to be present.
const DEFAULT_STYLE_ID: u32 = 201;

fn push_entry(out: &mut String, e: &Entry) {
    out.push_str("<cs:");
    out.push_str(e.name);
    if let Some(mods) = e.mods {
        out.push_str(r#" mods=""#);
        out.push_str(mods);
        out.push('"');
    }
    out.push('>');

    out.push_str(r#"<cs:lnRef idx="0"/>"#);
    if e.fill_ref_inner.is_empty() {
        out.push_str(r#"<cs:fillRef idx="0"/>"#);
    } else {
        out.push_str(r#"<cs:fillRef idx="0">"#);
        out.push_str(e.fill_ref_inner);
        out.push_str("</cs:fillRef>");
    }
    out.push_str(r#"<cs:effectRef idx="0"/>"#);
    out.push_str(r#"<cs:fontRef idx="minor">"#);
    out.push_str(e.font_ref_inner);
    out.push_str("</cs:fontRef>");

    if let Some(sp) = e.sp_pr_inner {
        out.push_str("<cs:spPr>");
        out.push_str(sp);
        out.push_str("</cs:spPr>");
    }
    if let Some(sz) = e.def_rpr_sz {
        out.push_str(r#"<cs:defRPr sz=""#);
        out.push_str(&sz.to_string());
        out.push_str(r#""/>"#);
    }

    out.push_str("</cs:");
    out.push_str(e.name);
    out.push('>');
}

/// Bytes of a default chart style part (`xl/charts/styleN.xml`).
pub fn default_chart_style_bytes() -> Vec<u8> {
    let mut out = String::with_capacity(4096);
    out.push_str(DECL);
    out.push_str(r#"<cs:chartStyle "#);
    out.push_str(NS);
    out.push_str(r#" id=""#);
    out.push_str(&DEFAULT_STYLE_ID.to_string());
    out.push_str(r#"">"#);

    for e in ENTRIES {
        push_entry(&mut out, e);
        if e.name == "dataPointMarker" {
            out.push_str(MARKER_LAYOUT);
        }
    }

    out.push_str("</cs:chartStyle>");
    out.into_bytes()
}

/// Bytes of a default chart colour style part (`xl/charts/colorsN.xml`).
///
/// The "colourful" palette: cycle through the six theme accents, then
/// through luminance variations of them, which is what Excel applies to
/// a new chart.
pub fn default_chart_color_style_bytes() -> Vec<u8> {
    let mut out = String::with_capacity(1024);
    out.push_str(DECL);
    out.push_str(r#"<cs:colorStyle "#);
    out.push_str(NS);
    out.push_str(r#" meth="cycle" id="10">"#);
    for i in 1..=6 {
        out.push_str(&format!(r#"<a:schemeClr val="accent{i}"/>"#));
    }
    out.push_str(concat!(
        r#"<cs:variation/>"#,
        r#"<cs:variation><a:lumMod val="60000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="80000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="50000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="70000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation>"#,
    ));
    out.push_str("</cs:colorStyle>");
    out.into_bytes()
}

/// Namespace both parts live in.
const CHART_STYLE_NS: &[u8] = b"http://schemas.microsoft.com/office/drawing/2012/chartStyle";

/// The `CT_ChartStyle` children with `minOccurs="1"`, in schema order.
/// `dataLabelCallout`, `dataPointMarkerLayout` and `extLst` are optional
/// and so are not required here.
const REQUIRED_ENTRIES: &[&str] = &[
    "axisTitle",
    "categoryAxis",
    "chartArea",
    "dataLabel",
    "dataPoint",
    "dataPoint3D",
    "dataPointLine",
    "dataPointMarker",
    "dataPointWireframe",
    "dataTable",
    "downBar",
    "dropLine",
    "errorBar",
    "floor",
    "gridlineMajor",
    "gridlineMinor",
    "hiLoLine",
    "leaderLine",
    "legend",
    "plotArea",
    "plotArea3D",
    "seriesAxis",
    "seriesLine",
    "title",
    "trendline",
    "trendlineLabel",
    "upBar",
    "valueAxis",
    "wall",
];

/// Root element name and its immediate children, in document order.
fn scan_part(bytes: &[u8], expected_root: &str) -> Result<(Vec<String>, bool), String> {
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;

    let mut reader = quick_xml::NsReader::from_reader(bytes);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut seen_root = false;
    let mut root_has_id = false;
    let mut children = Vec::new();

    loop {
        let (ns, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(|e| format!("not well-formed XML: {e}"))?;

        let (start, is_empty) = match event {
            Event::Start(ref e) => (Some(e.clone()), false),
            Event::Empty(ref e) => (Some(e.clone()), true),
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                buf.clear();
                continue;
            }
            Event::Eof => break,
            _ => {
                buf.clear();
                continue;
            }
        };

        if let Some(e) = start {
            let local = String::from_utf8_lossy(e.name().local_name().as_ref()).into_owned();
            if !seen_root {
                match ns {
                    ResolveResult::Unknown(prefix) => {
                        return Err(format!(
                            "namespace prefix `{}` on <{}> is not bound",
                            String::from_utf8_lossy(&prefix),
                            local
                        ))
                    }
                    ResolveResult::Bound(n) if n.as_ref() == CHART_STYLE_NS => {}
                    _ => {
                        return Err(format!(
                            "root <{local}> is not in the chartStyle namespace"
                        ))
                    }
                }
                if local != expected_root {
                    return Err(format!("root is <{local}>, expected <{expected_root}>"));
                }
                seen_root = true;
                root_has_id = e
                    .attributes()
                    .flatten()
                    .any(|a| a.key.local_name().as_ref() == b"id");
            } else if depth == 1 {
                children.push(local);
            }
            if !is_empty {
                depth += 1;
            }
        }
        buf.clear();
    }

    if !seen_root {
        return Err(format!("no <{expected_root}> root element"));
    }
    Ok((children, root_has_id))
}

/// Check that raw chart style bytes are a part Excel will accept.
///
/// Excel validates this part rather than repairing it: a missing `id`,
/// a missing required entry, or entries out of schema order all make it
/// refuse to open the workbook. Catching that here turns an unopenable
/// output file into an error naming the problem.
pub fn validate_chart_style_part(bytes: &[u8]) -> Result<(), String> {
    let (children, has_id) = scan_part(bytes, "chartStyle")?;

    if !has_id {
        // Optional per the schema, but Excel rejects the part without it.
        return Err("<cs:chartStyle> has no id attribute".to_string());
    }

    let mut next = 0usize;
    for child in &children {
        if next < REQUIRED_ENTRIES.len() && child == REQUIRED_ENTRIES[next] {
            next += 1;
        }
    }
    if next != REQUIRED_ENTRIES.len() {
        return Err(format!(
            "<cs:chartStyle> is missing required entry <cs:{}> or has entries out of schema order",
            REQUIRED_ENTRIES[next]
        ));
    }
    Ok(())
}

/// Check that raw chart colour style bytes are a part Excel will accept.
pub fn validate_chart_color_style_part(bytes: &[u8]) -> Result<(), String> {
    scan_part(bytes, "colorStyle").map(|_| ())
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The order and completeness of the sequence is what Excel checks.
    #[test]
    fn style_contains_every_required_entry_in_schema_order() {
        // CT_ChartStyle, [MS-ODRAWXML] §5.15.
        const SCHEMA_ORDER: &[&str] = &[
            "axisTitle",
            "categoryAxis",
            "chartArea",
            "dataLabel",
            "dataLabelCallout",
            "dataPoint",
            "dataPoint3D",
            "dataPointLine",
            "dataPointMarker",
            "dataPointMarkerLayout",
            "dataPointWireframe",
            "dataTable",
            "downBar",
            "dropLine",
            "errorBar",
            "floor",
            "gridlineMajor",
            "gridlineMinor",
            "hiLoLine",
            "leaderLine",
            "legend",
            "plotArea",
            "plotArea3D",
            "seriesAxis",
            "seriesLine",
            "title",
            "trendline",
            "trendlineLabel",
            "upBar",
            "valueAxis",
            "wall",
        ];

        let xml = String::from_utf8(default_chart_style_bytes()).unwrap();
        let mut cursor = 0usize;
        for name in SCHEMA_ORDER {
            // Boundary-terminated so `<cs:dataPoint` cannot satisfy the
            // search for `<cs:dataPoint3D` and vice versa.
            let hit = [">", " ", "/"]
                .iter()
                .filter_map(|end| xml[cursor..].find(&format!("<cs:{name}{end}")))
                .min()
                .unwrap_or_else(|| panic!("{name} missing from generated chart style"));
            cursor += hit + name.len();
        }
    }

    /// Excel requires `id` even though the schema marks it optional.
    #[test]
    fn style_carries_an_id_attribute() {
        let xml = String::from_utf8(default_chart_style_bytes()).unwrap();
        assert!(xml.contains(r#"<cs:chartStyle "#) && xml.contains(r#" id=""#));
    }

    /// Every `CT_StyleEntry` needs all four required reference children.
    #[test]
    fn every_style_entry_has_its_required_children() {
        let xml = String::from_utf8(default_chart_style_bytes()).unwrap();
        for e in ENTRIES {
            let start = xml
                .find(&format!("<cs:{}", e.name))
                .unwrap_or_else(|| panic!("{} missing", e.name));
            let end = xml
                .find(&format!("</cs:{}>", e.name))
                .unwrap_or_else(|| panic!("{} unterminated", e.name));
            let body = &xml[start..end];
            for required in ["<cs:lnRef", "<cs:fillRef", "<cs:effectRef", "<cs:fontRef"] {
                assert!(
                    body.contains(required),
                    "{} is missing {required}",
                    e.name
                );
            }
        }
    }

    #[test]
    fn both_parts_are_well_formed_xml() {
        for bytes in [default_chart_style_bytes(), default_chart_color_style_bytes()] {
            let mut reader = quick_xml::Reader::from_reader(&bytes[..]);
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(quick_xml::events::Event::Eof) => break,
                    Ok(_) => {}
                    Err(e) => panic!("generated part is not well formed: {e}"),
                }
                buf.clear();
            }
        }
    }

    /// The parts we generate must satisfy the checks we apply to
    /// parts supplied by callers.
    #[test]
    fn generated_parts_pass_their_own_validation() {
        validate_chart_style_part(&default_chart_style_bytes()).expect("generated style");
        validate_chart_color_style_part(&default_chart_color_style_bytes())
            .expect("generated colour style");
    }

    /// Each case below is a package Excel was observed to refuse.
    #[test]
    fn validation_rejects_parts_excel_refuses() {
        // Unbound `cs:` prefix: not a namespace-well-formed document.
        let err = validate_chart_style_part(br#"<cs:chartStyle/>"#).unwrap_err();
        assert!(err.contains("not bound"), "{err}");

        // Well formed and correctly bound, but no entries.
        let empty = format!(r#"<cs:chartStyle {NS} id="1"/>"#);
        let err = validate_chart_style_part(empty.as_bytes()).unwrap_err();
        assert!(err.contains("axisTitle"), "{err}");

        // Complete, but missing the id attribute Excel insists on.
        let mut body = String::new();
        for e in ENTRIES {
            push_entry(&mut body, e);
        }
        let no_id = format!(r#"<cs:chartStyle {NS}>{body}</cs:chartStyle>"#);
        let err = validate_chart_style_part(no_id.as_bytes()).unwrap_err();
        assert!(err.contains("id attribute"), "{err}");

        // Complete and identified, but in reverse order.
        let mut reversed = String::new();
        for e in ENTRIES.iter().rev() {
            push_entry(&mut reversed, e);
        }
        let out_of_order = format!(r#"<cs:chartStyle {NS} id="1">{reversed}</cs:chartStyle>"#);
        let err = validate_chart_style_part(out_of_order.as_bytes()).unwrap_err();
        assert!(err.contains("schema order"), "{err}");
    }

    #[test]
    fn validation_rejects_a_wrong_root_element() {
        let colors = default_chart_color_style_bytes();
        let err = validate_chart_style_part(&colors).unwrap_err();
        assert!(err.contains("expected <chartStyle>"), "{err}");

        let style = default_chart_style_bytes();
        let err = validate_chart_color_style_part(&style).unwrap_err();
        assert!(err.contains("expected <colorStyle>"), "{err}");
    }

    /// A dropped required entry must be named, not merely refused.
    #[test]
    fn validation_names_the_missing_entry() {
        let mut body = String::new();
        for e in ENTRIES.iter().filter(|e| e.name != "wall") {
            push_entry(&mut body, e);
        }
        let xml = format!(r#"<cs:chartStyle {NS} id="1">{body}</cs:chartStyle>"#);
        let err = validate_chart_style_part(xml.as_bytes()).unwrap_err();
        assert!(err.contains("wall"), "{err}");
    }

    #[test]
    fn color_style_cycles_the_six_theme_accents() {
        let xml = String::from_utf8(default_chart_color_style_bytes()).unwrap();
        assert!(xml.contains(r#"meth="cycle""#));
        for i in 1..=6 {
            assert!(xml.contains(&format!(r#"<a:schemeClr val="accent{i}"/>"#)));
        }
    }
}
