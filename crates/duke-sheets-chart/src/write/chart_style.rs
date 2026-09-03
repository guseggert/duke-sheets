//! Serializing the chart style and chart colour style parts.
//!
//! Order is load-bearing. Excel validates the style part rather than
//! repairing it: an entry out of the `CT_ChartStyle` sequence, or one
//! missing, and it refuses the workbook. The sequence is written from
//! one table so the order cannot drift from the parser's.

use std::io::Write;

use crate::chart_style::{
    ChartColorStyle, ChartColorStylePart, ChartStyle, ChartStylePart, FontReference, MarkerLayout,
    StyleEntry, StyleReference,
};

const NS: &str = concat!(
    r#"xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#
);

const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// The `CT_ChartStyle` sequence: element name, and how to reach that
/// entry. `dataPointMarkerLayout` is a marker layout rather than a style
/// entry and is handled where it sits in the order.
type EntryAccessor = fn(&ChartStyle) -> &StyleEntry;

/// Every required entry, in schema order, paired with the optional
/// element that follows it.
const SEQUENCE: &[(&str, EntryAccessor)] = &[
    ("axisTitle", |s| &s.axis_title),
    ("categoryAxis", |s| &s.category_axis),
    ("chartArea", |s| &s.chart_area),
    ("dataLabel", |s| &s.data_label),
    ("dataPoint", |s| &s.data_point),
    ("dataPoint3D", |s| &s.data_point_3d),
    ("dataPointLine", |s| &s.data_point_line),
    ("dataPointMarker", |s| &s.data_point_marker),
    ("dataPointWireframe", |s| &s.data_point_wireframe),
    ("dataTable", |s| &s.data_table),
    ("downBar", |s| &s.down_bar),
    ("dropLine", |s| &s.drop_line),
    ("errorBar", |s| &s.error_bar),
    ("floor", |s| &s.floor),
    ("gridlineMajor", |s| &s.gridline_major),
    ("gridlineMinor", |s| &s.gridline_minor),
    ("hiLoLine", |s| &s.hi_lo_line),
    ("leaderLine", |s| &s.leader_line),
    ("legend", |s| &s.legend),
    ("plotArea", |s| &s.plot_area),
    ("plotArea3D", |s| &s.plot_area_3d),
    ("seriesAxis", |s| &s.series_axis),
    ("seriesLine", |s| &s.series_line),
    ("title", |s| &s.title),
    ("trendline", |s| &s.trendline),
    ("trendlineLabel", |s| &s.trendline_label),
    ("upBar", |s| &s.up_bar),
    ("valueAxis", |s| &s.value_axis),
    ("wall", |s| &s.wall),
];

fn push_reference(out: &mut Vec<u8>, name: &str, reference: &StyleReference) {
    let _ = write!(out, r#"<cs:{name} idx="{}""#, reference.idx);
    match reference.color {
        Some(ref color) => {
            out.extend_from_slice(b">");
            out.extend_from_slice(color);
            let _ = write!(out, "</cs:{name}>");
        }
        None => out.extend_from_slice(b"/>"),
    }
}
fn push_font_reference(out: &mut Vec<u8>, reference: &FontReference) {
    let _ = write!(
        out,
        r#"<cs:fontRef idx="{}""#,
        reference.collection.as_str()
    );
    match reference.color {
        Some(ref color) => {
            out.extend_from_slice(b">");
            out.extend_from_slice(color);
            out.extend_from_slice(b"</cs:fontRef>");
        }
        None => out.extend_from_slice(b"/>"),
    }
}

/// `CT_StyleEntry` order: lnRef, lineWidthScale, fillRef, effectRef,
/// fontRef, spPr, defRPr, bodyPr, extLst.
fn push_entry(out: &mut Vec<u8>, name: &str, entry: &StyleEntry) {
    let _ = write!(out, "<cs:{name}");
    if let Some(ref mods) = entry.mods {
        let _ = write!(out, r#" mods="{mods}""#);
    }
    out.extend_from_slice(b">");

    push_reference(out, "lnRef", &entry.line_reference);
    if let Some(scale) = entry.line_width_scale {
        let _ = write!(out, r#"<cs:lineWidthScale>{scale}</cs:lineWidthScale>"#);
    }
    push_reference(out, "fillRef", &entry.fill_reference);
    push_reference(out, "effectRef", &entry.effect_reference);
    push_font_reference(out, &entry.font_reference);
    for raw in [
        entry.shape_properties.as_ref(),
        entry.default_run_properties.as_ref(),
        entry.body_properties.as_ref(),
        entry.extensions.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        out.extend_from_slice(raw);
    }

    let _ = write!(out, "</cs:{name}>");
}

fn push_marker_layout(out: &mut Vec<u8>, layout: &MarkerLayout) {
    out.extend_from_slice(b"<cs:dataPointMarkerLayout");
    if let Some(ref symbol) = layout.symbol {
        let _ = write!(out, r#" symbol="{symbol}""#);
    }
    if let Some(size) = layout.size {
        let _ = write!(out, r#" size="{size}""#);
    }
    out.extend_from_slice(b"/>");
}

/// Serialize a chart style part.
pub fn chart_style_bytes(style: &ChartStyle) -> Vec<u8> {
    let mut out = Vec::with_capacity(8 * 1024);
    out.extend_from_slice(DECL.as_bytes());
    let _ = write!(out, r#"<cs:chartStyle {NS} id="{}">"#, style.id);

    for (name, entry) in SEQUENCE {
        push_entry(&mut out, name, entry(style));
        // The two optional elements sit inside the sequence, not after
        // it: dataLabelCallout follows dataLabel, and the marker layout
        // follows dataPointMarker.
        match *name {
            "dataLabel" => {
                if let Some(ref callout) = style.data_label_callout {
                    push_entry(&mut out, "dataLabelCallout", callout);
                }
            }
            "dataPointMarker" => {
                if let Some(ref layout) = style.data_point_marker_layout {
                    push_marker_layout(&mut out, layout);
                }
            }
            _ => {}
        }
    }

    if let Some(ref extensions) = style.extensions {
        out.extend_from_slice(extensions);
    }
    out.extend_from_slice(b"</cs:chartStyle>");
    out
}

/// Serialize a chart colour style part.
pub fn chart_color_style_bytes(style: &ChartColorStyle) -> Vec<u8> {
    let mut out = Vec::with_capacity(1024);
    out.extend_from_slice(DECL.as_bytes());
    let _ = write!(out, r#"<cs:colorStyle {NS} meth="{}""#, style.method.as_str());
    if let Some(id) = style.id {
        let _ = write!(out, r#" id="{id}""#);
    }
    out.extend_from_slice(b">");
    for color in &style.colors {
        out.extend_from_slice(color);
    }
    for variation in &style.variations {
        out.extend_from_slice(variation);
    }
    out.extend_from_slice(b"</cs:colorStyle>");
    out
}

/// The bytes of a chart style part, whether modelled or kept as read.
pub fn chart_style_part_bytes(part: &ChartStylePart) -> Vec<u8> {
    match part {
        ChartStylePart::Typed(style) => chart_style_bytes(style),
        ChartStylePart::Raw(bytes) => bytes.clone(),
    }
}

/// The bytes of a chart colour style part, whether modelled or kept as
/// read.
pub fn chart_color_style_part_bytes(part: &ChartColorStylePart) -> Vec<u8> {
    match part {
        ChartColorStylePart::Typed(style) => chart_color_style_bytes(style),
        ChartColorStylePart::Raw(bytes) => bytes.clone(),
    }
}

#[cfg(all(test, feature = "parse"))]
mod tests {
    use super::*;
    use crate::chart_style::{ChartColorStyle, ChartStyle};
    use crate::parse::{parse_chart_color_style, parse_chart_style};

    /// The element names live in three places - the parser's sequence,
    /// the writer's, and the accessor the bindings surface entries by -
    /// and a drift between them would silently emit an entry out of
    /// order or label one as another.
    #[test]
    fn the_three_name_lists_agree() {
        let parse_order = crate::parse::chart_style_sequence();
        let required: Vec<&str> = parse_order
            .iter()
            .copied()
            .filter(|n| !matches!(*n, "dataLabelCallout" | "dataPointMarkerLayout"))
            .collect();

        let written: Vec<&str> = SEQUENCE.iter().map(|(name, _)| *name).collect();
        assert_eq!(written, required, "the writer's sequence");

        let style = ChartStyle::default();
        let accessed: Vec<&str> = crate::chart_style::entries_by_name(&style)
            .into_iter()
            .map(|(name, _)| name)
            .collect();
        let expected: Vec<&str> = parse_order
            .iter()
            .copied()
            .filter(|n| *n != "dataPointMarkerLayout")
            .collect();
        assert_eq!(accessed, expected, "the accessor the bindings use");
    }

    #[test]
    fn default_style_round_trips_through_the_model() {
        let style = ChartStyle::default();
        let bytes = chart_style_bytes(&style);
        let parsed = parse_chart_style(&bytes[..]).expect("the default must be modellable");
        assert_eq!(parsed, style);
    }

    #[test]
    fn default_color_style_round_trips_through_the_model() {
        let style = ChartColorStyle::default();
        let bytes = chart_color_style_bytes(&style);
        let parsed = parse_chart_color_style(&bytes[..]).expect("the default must be modellable");
        assert_eq!(parsed, style);
    }
}
