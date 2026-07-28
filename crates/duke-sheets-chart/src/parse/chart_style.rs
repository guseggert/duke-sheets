//! Reading the chart style and chart colour style parts.
//!
//! Strict on purpose. Excel validates the style part rather than
//! repairing it, so a part this parser accepts is one Excel will accept:
//! every required entry present, exactly once, in `CT_ChartStyle` order,
//! `id` present, nothing unrecognised. A part that fails any of those is
//! reported rather than coerced, and the caller keeps the bytes instead
//! ([`crate::chart_style::ChartStylePart::Raw`]), so reading a file stays
//! permissive without the model having to represent something Excel
//! would reject.

use std::io::BufRead;

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::NsReader;

use crate::chart_style::{
    ChartColorStyle, ChartStyle, ColorMethod, FontCollection, FontReference, MarkerLayout,
    StyleEntry, StyleReference,
};

const CHART_STYLE_NS: &[u8] = b"http://schemas.microsoft.com/office/drawing/2012/chartStyle";

/// Why a part could not be modelled.
pub type StyleParseError = String;

/// The `CT_ChartStyle` sequence. `dataLabelCallout` and
/// `dataPointMarkerLayout` are optional; every other name is required
/// exactly once, in this order.
const SEQUENCE: &[&str] = &[
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

/// The `CT_ChartStyle` sequence, for cross-checking the other places
/// these names are spelt out.
pub fn chart_style_sequence() -> &'static [&'static str] {
    SEQUENCE
}

fn optional(name: &str) -> bool {
    matches!(name, "dataLabelCallout" | "dataPointMarkerLayout")
}

/// One element of a part, as read: name, attributes, and either nested
/// elements or the bytes of the whole thing.
struct Element {
    name: String,
    /// Present when the element is in the chartStyle namespace.
    cs: bool,
    attrs: Vec<(String, String)>,
    children: Vec<Element>,
    /// The element as written, from its start tag to its end tag.
    raw: Vec<u8>,
}

impl Element {
    fn attr(&self, name: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.as_str())
    }

    fn child(&self, name: &str) -> Option<&Element> {
        self.children.iter().find(|child| child.name == name)
    }
}

/// Read a whole part into a tree, keeping each element exactly as it was
/// written, so the pieces this crate does not model can be replayed byte
/// for byte - including whether they were self-closing.
fn read_tree<R: BufRead>(reader: R) -> Result<Element, StyleParseError> {
    let mut reader = NsReader::from_reader(reader);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut stack: Vec<Element> = Vec::new();
    let mut root: Option<Element> = None;

    /// Serialize one event exactly as it appeared.
    fn as_written(event: &Event) -> Vec<u8> {
        let mut sink = quick_xml::Writer::new(Vec::new());
        let _ = sink.write_event(event.clone());
        sink.into_inner()
    }

    loop {
        let (ns, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(|e| format!("not well-formed XML: {e}"))?;

        match event {
            Event::Eof => break,
            Event::End(ref e) => {
                let mut done = stack.pop().ok_or("stray end tag")?;
                done.raw.extend_from_slice(&as_written(&Event::End(e.clone())));
                match stack.last_mut() {
                    Some(parent) => {
                        parent.raw.extend_from_slice(&done.raw);
                        parent.children.push(done);
                    }
                    None => root = Some(done),
                }
            }
            Event::Start(ref e) | Event::Empty(ref e) => {
                let empty = matches!(event, Event::Empty(_));
                let name = String::from_utf8_lossy(e.name().local_name().as_ref()).into_owned();
                let cs = match ns {
                    ResolveResult::Unknown(prefix) => {
                        return Err(format!(
                            "namespace prefix `{}` on <{name}> is not bound",
                            String::from_utf8_lossy(&prefix)
                        ))
                    }
                    ResolveResult::Bound(bound) => bound.as_ref() == CHART_STYLE_NS,
                    ResolveResult::Unbound => false,
                };
                let attrs = e
                    .attributes()
                    .flatten()
                    .map(|a| {
                        (
                            String::from_utf8_lossy(a.key.local_name().as_ref()).into_owned(),
                            a.unescape_value().unwrap_or_default().to_string(),
                        )
                    })
                    .collect();
                let element = Element {
                    name,
                    cs,
                    attrs,
                    children: Vec::new(),
                    raw: as_written(&event),
                };
                if empty {
                    match stack.last_mut() {
                        Some(parent) => {
                            parent.raw.extend_from_slice(&element.raw);
                            parent.children.push(element);
                        }
                        None => root = Some(element),
                    }
                } else {
                    stack.push(element);
                }
            }
            ref other => {
                // Text, comments and CDATA belong to whichever element is
                // open, and are only ever replayed.
                if let Some(open) = stack.last_mut() {
                    open.raw.extend_from_slice(&as_written(other));
                }
            }
        }
        buf.clear();
    }

    root.ok_or_else(|| "no root element".to_string())
}

fn parse_reference(entry: &Element, name: &str) -> Result<StyleReference, StyleParseError> {
    let element = entry
        .child(name)
        .ok_or_else(|| format!("<cs:{}> has no <cs:{name}>", entry.name))?;
    let idx = element
        .attr("idx")
        .ok_or_else(|| format!("<cs:{name}> has no idx"))?
        .parse::<u32>()
        .map_err(|_| format!("<cs:{name}> idx is not a number"))?;
    Ok(StyleReference {
        idx,
        color: element.children.first().map(|child| child.raw.clone()),
    })
}

fn parse_entry(element: &Element) -> Result<StyleEntry, StyleParseError> {
    let font = element
        .child("fontRef")
        .ok_or_else(|| format!("<cs:{}> has no <cs:fontRef>", element.name))?;
    let collection = font
        .attr("idx")
        .and_then(FontCollection::from_str_value)
        .ok_or_else(|| format!("<cs:{}> has an unknown fontRef idx", element.name))?;

    let raw_child = |name: &str| {
        element
            .children
            .iter()
            .find(|child| !child.cs && child.name == name)
            .map(|child| child.raw.clone())
    };

    Ok(StyleEntry {
        line_reference: parse_reference(element, "lnRef")?,
        line_width_scale: element
            .child("lineWidthScale")
            .and_then(|scale| String::from_utf8_lossy(&scale.raw).parse::<f64>().ok()),
        fill_reference: parse_reference(element, "fillRef")?,
        effect_reference: parse_reference(element, "effectRef")?,
        font_reference: FontReference {
            collection,
            color: font.children.first().map(|child| child.raw.clone()),
        },
        shape_properties: raw_child("spPr"),
        default_run_properties: raw_child("defRPr"),
        body_properties: raw_child("bodyPr"),
        extensions: element
            .children
            .iter()
            .find(|child| child.cs && child.name == "extLst")
            .map(|child| child.raw.clone()),
        mods: element.attr("mods").map(str::to_string),
    })
}

/// Read a chart style part, or say why it cannot be modelled.
pub fn parse_chart_style<R: BufRead>(reader: R) -> Result<ChartStyle, StyleParseError> {
    let root = read_tree(reader)?;
    if !root.cs || root.name != "chartStyle" {
        return Err(format!("root is <{}>, expected <cs:chartStyle>", root.name));
    }
    let id = root
        .attr("id")
        // Optional per the schema, but Excel refuses the part without it.
        .ok_or("<cs:chartStyle> has no id attribute")?
        .parse::<u32>()
        .map_err(|_| "<cs:chartStyle> id is not a number".to_string())?;

    // Walk the children against the sequence: each must be the next name
    // expected, optional ones may be skipped, nothing else is allowed.
    let mut entries: Vec<(&str, &Element)> = Vec::new();
    let mut next = 0usize;
    for child in &root.children {
        if child.cs && child.name == "extLst" {
            continue;
        }
        if !child.cs {
            return Err(format!("<{}> is not a chartStyle element", child.name));
        }
        let found = SEQUENCE[next..]
            .iter()
            .position(|name| *name == child.name)
            .ok_or_else(|| {
                format!(
                    "<cs:{}> is unexpected here; the sequence expects <cs:{}> next",
                    child.name,
                    SEQUENCE.get(next).copied().unwrap_or("extLst")
                )
            })?;
        for skipped in &SEQUENCE[next..next + found] {
            if !optional(skipped) {
                return Err(format!("<cs:chartStyle> is missing <cs:{skipped}>"));
            }
        }
        entries.push((SEQUENCE[next + found], child));
        next += found + 1;
    }
    for remaining in &SEQUENCE[next..] {
        if !optional(remaining) {
            return Err(format!("<cs:chartStyle> is missing <cs:{remaining}>"));
        }
    }

    let mut style = ChartStyle {
        id,
        extensions: root
            .children
            .iter()
            .find(|child| child.cs && child.name == "extLst")
            .map(|child| child.raw.clone()),
        ..ChartStyle::default()
    };
    // Clear the defaults the entries below replace, so a part that omits
    // an optional element does not inherit one.
    style.data_label_callout = None;
    style.data_point_marker_layout = None;

    for (name, element) in entries {
        if name == "dataPointMarkerLayout" {
            style.data_point_marker_layout = Some(MarkerLayout {
                symbol: element.attr("symbol").map(str::to_string),
                size: element.attr("size").and_then(|s| s.parse().ok()),
            });
            continue;
        }
        let entry = parse_entry(element)?;
        match name {
            "axisTitle" => style.axis_title = entry,
            "categoryAxis" => style.category_axis = entry,
            "chartArea" => style.chart_area = entry,
            "dataLabel" => style.data_label = entry,
            "dataLabelCallout" => style.data_label_callout = Some(entry),
            "dataPoint" => style.data_point = entry,
            "dataPoint3D" => style.data_point_3d = entry,
            "dataPointLine" => style.data_point_line = entry,
            "dataPointMarker" => style.data_point_marker = entry,
            "dataPointWireframe" => style.data_point_wireframe = entry,
            "dataTable" => style.data_table = entry,
            "downBar" => style.down_bar = entry,
            "dropLine" => style.drop_line = entry,
            "errorBar" => style.error_bar = entry,
            "floor" => style.floor = entry,
            "gridlineMajor" => style.gridline_major = entry,
            "gridlineMinor" => style.gridline_minor = entry,
            "hiLoLine" => style.hi_lo_line = entry,
            "leaderLine" => style.leader_line = entry,
            "legend" => style.legend = entry,
            "plotArea" => style.plot_area = entry,
            "plotArea3D" => style.plot_area_3d = entry,
            "seriesAxis" => style.series_axis = entry,
            "seriesLine" => style.series_line = entry,
            "title" => style.title = entry,
            "trendline" => style.trendline = entry,
            "trendlineLabel" => style.trendline_label = entry,
            "upBar" => style.up_bar = entry,
            "valueAxis" => style.value_axis = entry,
            "wall" => style.wall = entry,
            other => return Err(format!("<cs:{other}> is not a chartStyle entry")),
        }
    }
    Ok(style)
}

/// Read a chart colour style part, or say why it cannot be modelled.
pub fn parse_chart_color_style<R: BufRead>(reader: R) -> Result<ChartColorStyle, StyleParseError> {
    const COLOR_CHOICES: &[&str] = &[
        "scrgbClr", "srgbClr", "hslClr", "sysClr", "schemeClr", "prstClr",
    ];

    let root = read_tree(reader)?;
    if !root.cs || root.name != "colorStyle" {
        return Err(format!("root is <{}>, expected <cs:colorStyle>", root.name));
    }
    // `meth` is use="required" in CT_ColorStyle.
    let method = ColorMethod::from_str_value(
        root.attr("meth")
            .ok_or("<cs:colorStyle> has no meth attribute")?,
    );

    let mut colors = Vec::new();
    let mut variations = Vec::new();
    for child in &root.children {
        if child.cs && child.name == "variation" {
            variations.push(child.raw.clone());
        } else if !child.cs && COLOR_CHOICES.contains(&child.name.as_str()) {
            colors.push(child.raw.clone());
        } else {
            return Err(format!("<{}> is not a colour or a variation", child.name));
        }
    }
    if colors.is_empty() {
        // EG_ColorChoice is minOccurs="1".
        return Err("<cs:colorStyle> has no colour".to_string());
    }

    Ok(ChartColorStyle {
        method,
        id: root.attr("id").and_then(|id| id.parse().ok()),
        colors,
        variations,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart_style::ChartStyle;

    const NS: &str = concat!(
        r#"xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#
    );

    /// The four references CT_StyleEntry requires.
    fn refs() -> String {
        concat!(
            r#"<cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/>"#,
            r#"<cs:fontRef idx="minor"/>"#
        )
        .to_string()
    }

    fn entries(names: &[&str]) -> String {
        names
            .iter()
            .map(|n| format!("<cs:{n}>{}</cs:{n}>", refs()))
            .collect()
    }

    /// Every required name, in schema order.
    fn required() -> Vec<&'static str> {
        SEQUENCE.iter().copied().filter(|n| !optional(n)).collect()
    }

    fn style_doc(body: &str, id: Option<&str>) -> String {
        let id = id.map(|v| format!(r#" id="{v}""#)).unwrap_or_default();
        format!(r#"<cs:chartStyle {NS}{id}>{body}</cs:chartStyle>"#)
    }

    #[test]
    fn a_complete_part_is_modelled() {
        let doc = style_doc(&entries(&required()), Some("272"));
        let style = parse_chart_style(doc.as_bytes()).expect("complete part");
        assert_eq!(style.id, 272);
        // The optional elements are absent, not defaulted in.
        assert!(style.data_label_callout.is_none());
        assert!(style.data_point_marker_layout.is_none());
    }

    /// Excel requires `id` even though the schema marks it optional.
    #[test]
    fn a_part_without_an_id_is_not_modelled() {
        let doc = style_doc(&entries(&required()), None);
        let error = parse_chart_style(doc.as_bytes()).unwrap_err();
        assert!(error.contains("id attribute"), "{error}");
    }

    #[test]
    fn a_part_missing_a_required_entry_is_not_modelled() {
        let mut names = required();
        names.retain(|n| *n != "wall");
        let doc = style_doc(&entries(&names), Some("1"));
        let error = parse_chart_style(doc.as_bytes()).unwrap_err();
        assert!(error.contains("wall"), "{error}");
    }

    #[test]
    fn a_part_with_entries_out_of_order_is_not_modelled() {
        let mut names = required();
        names.reverse();
        let doc = style_doc(&entries(&names), Some("1"));
        let error = parse_chart_style(doc.as_bytes()).unwrap_err();
        // Reported either as the entry that turned up early or as the one
        // it skipped past; both name the sequence being violated.
        assert!(
            error.contains("unexpected here") || error.contains("is missing"),
            "{error}"
        );
    }

    /// Each required entry appears exactly once. A duplicate is a
    /// backwards step in the sequence, and taking the second silently
    /// would drop whatever the first said.
    #[test]
    fn a_part_with_a_duplicated_entry_is_not_modelled() {
        let mut names = required();
        names.insert(3, "axisTitle");
        let doc = style_doc(&entries(&names), Some("1"));
        let error = parse_chart_style(doc.as_bytes()).unwrap_err();
        assert!(error.contains("axisTitle"), "{error}");
    }

    /// An entry that turns up after one the sequence puts later is out of
    /// order even though every name is present exactly once.
    #[test]
    fn a_part_with_two_entries_transposed_is_not_modelled() {
        let mut names = required();
        let last = names.len() - 1;
        names.swap(last - 1, last);
        let doc = style_doc(&entries(&names), Some("1"));
        let error = parse_chart_style(doc.as_bytes()).unwrap_err();
        assert!(
            error.contains("unexpected here") || error.contains("is missing"),
            "{error}"
        );
    }

    #[test]
    fn a_part_whose_entries_lack_their_references_is_not_modelled() {
        let bare: String = required().iter().map(|n| format!("<cs:{n}/>")).collect();
        let doc = style_doc(&bare, Some("1"));
        let error = parse_chart_style(doc.as_bytes()).unwrap_err();
        assert!(error.contains("Ref>"), "must name the missing reference: {error}");
    }

    #[test]
    fn a_part_with_an_unbound_prefix_is_not_modelled() {
        let error = parse_chart_style(&b"<cs:chartStyle/>"[..]).unwrap_err();
        assert!(error.contains("not bound"), "{error}");
    }

    #[test]
    fn a_part_with_the_wrong_root_is_not_modelled() {
        let doc = format!(
            r#"<cs:colorStyle {NS} meth="cycle"><a:schemeClr val="accent1"/></cs:colorStyle>"#
        );
        let error = parse_chart_style(doc.as_bytes()).unwrap_err();
        assert!(error.contains("expected <cs:chartStyle>"), "{error}");

        let style = style_doc(&entries(&required()), Some("1"));
        let error = parse_chart_color_style(style.as_bytes()).unwrap_err();
        assert!(error.contains("expected <cs:colorStyle>"), "{error}");
    }

    #[test]
    fn a_part_with_a_foreign_entry_is_not_modelled() {
        let body = format!(
            r#"<x:axisTitle xmlns:x="urn:not-chartstyle">{}</x:axisTitle>"#,
            refs()
        );
        let doc = style_doc(&body, Some("1"));
        let error = parse_chart_style(doc.as_bytes()).unwrap_err();
        assert!(error.contains("not a chartStyle element"), "{error}");
    }

    /// The optional elements sit inside the sequence, so a part carrying
    /// them is still in order.
    #[test]
    fn the_optional_elements_are_accepted_in_their_place() {
        let mut body = String::new();
        for name in SEQUENCE {
            if *name == "dataPointMarkerLayout" {
                body.push_str(r#"<cs:dataPointMarkerLayout symbol="circle" size="7"/>"#);
            } else {
                body.push_str(&format!("<cs:{name}>{}</cs:{name}>", refs()));
            }
        }
        let doc = style_doc(&body, Some("1"));
        let style = parse_chart_style(doc.as_bytes()).expect("in order");
        assert!(style.data_label_callout.is_some());
        assert_eq!(
            style.data_point_marker_layout.as_ref().and_then(|m| m.size),
            Some(7)
        );
    }

    #[test]
    fn a_colour_style_needs_a_method_and_a_colour() {
        let no_meth = format!(r#"<cs:colorStyle {NS}><a:schemeClr val="accent1"/></cs:colorStyle>"#);
        let error = parse_chart_color_style(no_meth.as_bytes()).unwrap_err();
        assert!(error.contains("meth"), "{error}");

        let no_colors = format!(r#"<cs:colorStyle {NS} meth="cycle"><cs:variation/></cs:colorStyle>"#);
        let error = parse_chart_color_style(no_colors.as_bytes()).unwrap_err();
        assert!(error.contains("no colour"), "{error}");
    }

    /// What the entries hold is DrawingML, kept as the bytes it was
    /// written as, self-closing forms included.
    #[test]
    fn an_entry_keeps_its_payload_verbatim() {
        let mut names = required();
        names.retain(|n| *n != "chartArea");
        let with_payload = r#"<cs:chartArea mods="allowNoFillOverride"><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="major"><a:schemeClr val="dk1"/></cs:fontRef><a:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:spPr><a:defRPr sz="1000"/></cs:chartArea>"#;
        // chartArea is third in the sequence.
        let body = format!(
            "{}{}{}",
            entries(&names[..2]),
            with_payload,
            entries(&names[2..])
        );
        let doc = style_doc(&body, Some("1"));
        let style = parse_chart_style(doc.as_bytes()).expect("payload part");

        let area = &style.chart_area;
        assert_eq!(area.mods.as_deref(), Some("allowNoFillOverride"));
        assert_eq!(area.line_reference.idx, 1);
        assert_eq!(
            area.line_reference.color.as_deref().map(String::from_utf8_lossy),
            Some(r#"<a:schemeClr val="tx1"/>"#.into())
        );
        assert_eq!(area.font_reference.collection, FontCollection::Major);
        assert_eq!(
            area.shape_properties.as_deref().map(String::from_utf8_lossy),
            Some(r#"<a:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:spPr>"#.into())
        );
        assert_eq!(
            area.default_run_properties
                .as_deref()
                .map(String::from_utf8_lossy),
            Some(r#"<a:defRPr sz="1000"/>"#.into())
        );
    }

    /// The default is what a chart gets when a file had no style, and it
    /// has to be a part Excel accepts, so it must satisfy the same rules.
    #[test]
    fn the_default_satisfies_the_rules_a_read_part_must() {
        let bytes = crate::write::chart_style_bytes(&ChartStyle::default());
        parse_chart_style(&bytes[..]).expect("the generated default must be modellable");
    }
}
