#!/usr/bin/env python3
"""Fetch Open-XML-SDK schema JSONs and generate docs/CHART_SUPPORT.md.

Every checklist item includes the source schema filename and the
index into its Types array so we can trace back to the spec later.
"""

import json
import urllib.request
import sys
from collections import defaultdict

BASE = "https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/schemas/"

SCHEMAS = [
    # (short_label, filename)
    ("c", "schemas_openxmlformats_org_drawingml_2006_chart.json"),
    ("xdr", "schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json"),
    ("a", "schemas_openxmlformats_org_drawingml_2006_main.json"),
    ("c14", "schemas_microsoft_com_office_drawing_2007_8_2_chart.json"),
    ("c15", "schemas_microsoft_com_office_drawing_2012_chart.json"),
    ("c16", "schemas_microsoft_com_office_drawing_2014_chart.json"),
    ("cx", "schemas_microsoft_com_office_drawing_2014_chartex.json"),
    ("cs", "schemas_microsoft_com_office_drawing_2012_chartStyle.json"),
    ("c16r3", "schemas_microsoft_com_office_drawing_2017_03_chart.json"),
]

# Only include a: types that are actually referenced by chart/drawing
# elements. We collect these by scanning children/attributes of c: and
# xdr: types.
A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Categories for c: namespace types (matched by XSD type prefix)
CHART_TYPE_PREFIXES = [
    "CT_BarChart",
    "CT_Bar3DChart",
    "CT_LineChart",
    "CT_Line3DChart",
    "CT_PieChart",
    "CT_Pie3DChart",
    "CT_AreaChart",
    "CT_Area3DChart",
    "CT_ScatterChart",
    "CT_BubbleChart",
    "CT_RadarChart",
    "CT_DoughnutChart",
    "CT_StockChart",
    "CT_SurfaceChart",
    "CT_Surface3DChart",
    "CT_OfPieChart",
]

SERIES_PREFIXES = [
    "CT_BarSer",
    "CT_LineSer",
    "CT_PieSer",
    "CT_AreaSer",
    "CT_ScatterSer",
    "CT_BubbleSer",
    "CT_RadarSer",
    "CT_SurfaceSer",
]

AXIS_PREFIXES = [
    "CT_CatAx",
    "CT_ValAx",
    "CT_DateAx",
    "CT_SerAx",
    "CT_Scaling",
    "CT_AxisUnit",
]

DATA_REF_PREFIXES = [
    "CT_NumRef",
    "CT_StrRef",
    "CT_MultiLvlStrRef",
    "CT_NumData",
    "CT_StrData",
    "CT_NumDataSource",
    "CT_AxDataSource",
]

IMPLEMENTED_CLASSNAMES = {
    # xdr: SpreadsheetDrawing
    "WorksheetDrawing", "TwoCellAnchor", "OneCellAnchor",
    "FromMarker", "ToMarker",
    "ColumnId", "ColumnOffset", "RowId", "RowOffset",
    "GraphicFrame", "NonVisualGraphicFrameProperties",
    "NonVisualDrawingProperties", "NonVisualGraphicFrameDrawingProperties",
    "Transform", "ClientData",
    # c: chart structure
    "ChartSpace", "Chart", "PlotArea",
    "View3D", "RotateX", "RotateY", "DepthPercent",
    "HeightPercent", "Perspective",
    # c: chart types
    "BarChart", "Bar3DChart", "LineChart", "Line3DChart",
    "PieChart", "Pie3DChart", "DoughnutChart",
    "AreaChart", "Area3DChart",
    "ScatterChart", "BubbleChart",
    "RadarChart", "StockChart",
    "SurfaceChart", "Surface3DChart",
    "BarDirection", "Grouping", "BarGrouping",
    "ScatterStyle", "RadarStyle",
    # c: series
    "BarChartSeries", "LineChartSeries", "PieChartSeries",
    "AreaChartSeries", "ScatterChartSeries", "BubbleChartSeries",
    "RadarChartSeries", "SurfaceChartSeries",
    "Index", "Order", "SeriesText", "NumericValue",
    "Explosion",
    # c: data references
    "NumberReference", "StringReference",
    "NumberingCache", "NumberData",
    "StringCache", "StringData",
    "Formula", "NumericPoint", "StringPoint", "PointCount",
    "AxisDataSource", "NumberDataSource", "BubbleSize",
    # c: axes
    "CategoryAxis", "ValueAxis", "DateAxis", "SeriesAxis",
    "AxisId", "Scaling", "Orientation",
    "MinAxisValue", "MaxAxisValue",
    "Delete", "AxisPosition",
    "MajorGridlines", "MinorGridlines",
    "Title", "NumberingFormat",
    "MajorTickMark", "MinorTickMark", "TickLabelPosition",
    "CrossingAxis", "Crosses", "CrossBetween",
    "MajorUnit", "MinorUnit",
    # c: legend & title
    "Legend", "LegendPosition", "Overlay",
    "ChartText", "RichText", "SeriesText",
    # c: data labels & data points
    "DataLabels", "DataLabelPosition",
    "ShowLegendKey", "ShowValue", "ShowCategoryName",
    "ShowSeriesName", "ShowPercent", "ShowBubbleSize",
    "Separator", "ShowLeaderLines",
    "DataPoint",
    # c: trendlines & error bars
    "Trendline", "TrendlineType", "TrendlineName",
    "Period", "Forward", "Backward",
    "Intercept", "DisplayRSquaredValue", "DisplayEquation",
    "ErrorBars", "ErrorDirection", "ErrorBarType",
    "ErrorValueType", "ErrorBarValue", "NoEndCap",
    # c: formatting
    "Marker", "MarkerStyle", "MarkerSize",
    "ChartShapeProperties",
    # c: configuration
    "Layout", "ManualLayout",
    "Left", "Top", "Width", "Height",
    "DisplayBlanksAs", "DataTable",
    "ShowHorizontalBorder", "ShowVerticalBorder",
    "ShowOutlineBorder", "ShowKeys",
    "PlotVisibleOnly", "GapWidth", "Overlap",
    "VaryColors", "FirstSliceAngle", "HoleSize",
    "BubbleScale", "ShowNegativeBubbles", "Wireframe",
    "AutoTitleDeleted", "RoundedCorners",
    "ShowDataLabelsOverMaximum", "RightAngleAxes",
    "Smooth", "InvertIfNegative", "ShowMarker",
    # c: chart lines & up-down bars
    "DropLines", "HighLowLines", "SeriesLines",
    "UpDownBars", "UpBars", "DownBars",
    "LeaderLines",
    # c: extension lists
    "ExtensionList", "Extension",
    # a: DrawingML formatting subset
    "NoFill", "SolidFill", "SolidColorFillProperties",
    "RgbColorModelHex", "Outline", "LineProperties",
    "PresetDash",
    # xdr: additional anchors
    "AbsoluteAnchor", "Position", "Extent",
}

IMPLEMENTED_TAGS = {
    # xdr: SpreadsheetDrawing
    "xdr:wsDr", "xdr:twoCellAnchor", "xdr:oneCellAnchor",
    "xdr:from", "xdr:to",
    "xdr:col", "xdr:colOff", "xdr:row", "xdr:rowOff",
    "xdr:graphicFrame", "xdr:nvGraphicFramePr",
    "xdr:cNvPr", "xdr:cNvGraphicFramePr",
    "xdr:xfrm", "xdr:clientData",
    # c: chart structure
    "c:chartSpace", "c:chart", "c:plotArea",
    "c:view3D", "c:rotX", "c:rotY", "c:depthPercent",
    "c:hPercent", "c:perspective",
    # c: chart types
    "c:barChart", "c:bar3DChart",
    "c:lineChart", "c:line3DChart",
    "c:pieChart", "c:pie3DChart", "c:doughnutChart",
    "c:areaChart", "c:area3DChart",
    "c:scatterChart", "c:bubbleChart",
    "c:radarChart", "c:stockChart",
    "c:surfaceChart", "c:surface3DChart",
    "c:barDir", "c:grouping", "c:barGrouping",
    "c:scatterStyle", "c:radarStyle",
    # c: series & common children
    "c:ser",
    "c:idx", "c:order", "c:tx",
    "c:cat", "c:val", "c:xVal", "c:yVal",
    "c:explosion", "c:smooth", "c:invertIfNegative",
    "c:marker", "c:bubbleSize", "c:v",
    # c: data references
    "c:numRef", "c:strRef",
    "c:numCache", "c:numLit",
    "c:strCache", "c:strLit",
    "c:f", "c:pt", "c:ptCount",
    # c: axes
    "c:catAx", "c:valAx", "c:dateAx", "c:serAx",
    "c:axId", "c:scaling",
    "c:orientation", "c:min", "c:max",
    "c:delete", "c:axPos",
    "c:majorGridlines", "c:minorGridlines",
    "c:title", "c:numFmt",
    "c:majorTickMark", "c:minorTickMark", "c:tickLblPos",
    "c:crossAx", "c:crosses", "c:crossBetween",
    "c:majorUnit", "c:minorUnit",
    # c: legend & title
    "c:legend", "c:legendPos", "c:overlay",
    "c:rich",
    # c: data labels & data points
    "c:dLbls", "c:dLblPos",
    "c:showLegendKey", "c:showVal", "c:showCatName",
    "c:showSerName", "c:showPercent", "c:showBubbleSize",
    "c:separator", "c:showLeaderLines",
    "c:dPt",
    # c: trendlines & error bars
    "c:trendline", "c:trendlineType",
    "c:period", "c:forward", "c:backward",
    "c:intercept", "c:dispRSqr", "c:dispEq",
    "c:errBars", "c:errDir", "c:errBarType", "c:errValType",
    "c:noEndCap",
    # c: formatting
    "c:symbol", "c:size", "c:spPr",
    # c: configuration
    "c:layout", "c:manualLayout",
    "c:x", "c:y", "c:w", "c:h",
    "c:dispBlanksAs", "c:dTable",
    "c:showHorzBorder", "c:showVertBorder",
    "c:showOutline", "c:showKeys",
    "c:plotVisOnly", "c:gapWidth", "c:overlap",
    "c:varyColors", "c:firstSliceAng", "c:holeSize",
    "c:bubbleScale", "c:showNegBubbles", "c:wireframe",
    "c:autoTitleDeleted", "c:roundedCorners",
    "c:showDLblsOverMax", "c:rAngAx",
    # c: chart lines & up-down bars
    "c:dropLines", "c:hiLowLines", "c:serLines",
    "c:upDownBars", "c:upBars", "c:downBars",
    "c:leaderLines",
    # c: extension lists
    "c:extLst", "c:ext",
    # a: DrawingML formatting subset
    "a:noFill", "a:solidFill", "a:srgbClr",
    "a:ln", "a:prstDash",
    # xdr: additional anchors
    "xdr:absoluteAnchor", "xdr:pos", "xdr:ext",
}


def fetch_schema(filename):
    url = BASE + filename
    print(f"  Fetching {filename}...", file=sys.stderr)
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())


def xsd_type(name_field):
    """Extract the XSD complex type name from the Name field.

    Name looks like 'c:CT_BarChart/c:barChart' or 'c:CT_Boolean/'.
    Returns 'CT_BarChart'.
    """
    before_slash = name_field.split("/")[0]
    if ":" in before_slash:
        return before_slash.split(":", 1)[1]
    return before_slash


def xml_tag(name_field):
    """Extract the XML element tag from the Name field.

    Returns e.g. 'c:barChart' or '' for abstract bases.
    """
    parts = name_field.split("/", 1)
    return parts[1] if len(parts) > 1 else ""


def ns_prefix(name_field):
    """Extract the namespace prefix from the Name field."""
    tag = xml_tag(name_field) or name_field.split("/")[0]
    if ":" in tag:
        return tag.split(":")[0]
    return ""


def format_item(t, idx, schema_file):
    """Format a single type as a markdown checklist line."""
    name = t.get("Name", "")
    cls = t.get("ClassName", "")
    summary = t.get("Summary", "").rstrip(".")
    tag = xml_tag(name)
    is_abstract = t.get("IsAbstract", False)
    src = f"{schema_file}#Types[{idx}]"

    if is_abstract:
        return f"  - _abstract base: `{cls}`_ — {summary} (`{src}`)"

    implemented = cls in IMPLEMENTED_CLASSNAMES or tag in IMPLEMENTED_TAGS
    checkbox = "[x]" if implemented else "[ ]"
    tag_display = f"`{tag}`" if tag else f"`{xsd_type(name)}`"
    return f"- {checkbox} {tag_display} — **{cls}**: {summary} (`{src}`)"


def extract_enum_refs(types_list):
    """Extract unique enum type names from attribute Type fields."""
    enums = set()
    for t in types_list:
        for attr in t.get("Attributes", []):
            atype = attr.get("Type", "")
            if "EnumValue<" in atype:
                enum_name = atype.split("<", 1)[1].rstrip(">")
                # Take just the short name
                short = enum_name.rsplit(".", 1)[-1]
                enums.add(short)
    return sorted(enums)


def collect_a_refs(all_types_by_ns):
    """Find a: type names referenced by c: and xdr: types."""
    refs = set()
    for ns_label in ("c", "xdr", "c14", "c15", "c16", "cx", "cs", "c16r3"):
        for t in all_types_by_ns.get(ns_label, []):
            for child in t.get("Children", []):
                child_name = child.get("Name", "")
                if child_name.startswith("a:"):
                    refs.add(child_name.split("/")[0])  # e.g. "a:CT_ShapeProperties"
    return refs


def categorize_chart_type(t):
    """Categorize a c: namespace type into a section."""
    name = t.get("Name", "")
    xtype = xsd_type(name)

    for p in CHART_TYPE_PREFIXES:
        if xtype.startswith(p):
            return "chart_types"
    for p in SERIES_PREFIXES:
        if xtype.startswith(p):
            return "series"
    for p in AXIS_PREFIXES:
        if xtype.startswith(p):
            return "axes"
    for p in DATA_REF_PREFIXES:
        if xtype.startswith(p):
            return "data_refs"

    cls = t.get("ClassName", "")

    if xtype in ("CT_DLbls", "CT_DLbl", "CT_DLblPos") or "DataLabel" in cls:
        return "data_labels"
    if xtype.startswith("CT_DPt") or cls == "DataPoint":
        return "data_labels"

    if xtype.startswith("CT_Trendline") or "Trendline" in cls:
        return "trendlines"
    if xtype.startswith("CT_ErrBars") or "ErrorBar" in cls:
        return "trendlines"

    if xtype in ("CT_Legend", "CT_LegendEntry", "CT_LegendPos"):
        return "legend_title"
    if xtype in ("CT_Title", "CT_Tx", "CT_SerTx"):
        return "legend_title"
    if "Legend" in cls or cls in ("ChartText", "RichText", "SeriesText"):
        return "legend_title"

    if xtype.startswith("CT_Boolean") or (t.get("BaseClass", "") == "BooleanType"):
        return "booleans"

    if xtype.startswith("CT_UnsignedInt") or (
        t.get("BaseClass", "") == "UnsignedIntegerType"
    ):
        return "unsigned_ints"

    if xtype.startswith("CT_ChartLines") or (
        t.get("BaseClass", "") == "ChartLinesType"
    ):
        return "chart_config"

    if xtype.startswith("CT_Marker") and not xtype.startswith("CT_MarkerStyle"):
        return "formatting"
    if "ShapeProperties" in cls or "TextProperties" in cls:
        return "formatting"
    if "PictureOptions" in cls or "MarkerS" in xtype:
        return "formatting"

    if xtype in ("CT_Layout", "CT_ManualLayout", "CT_LayoutTarget", "CT_LayoutMode"):
        return "chart_config"
    if xtype.startswith("CT_View3D") or xtype in (
        "CT_RotX",
        "CT_RotY",
        "CT_DepthPercent",
        "CT_HPercent",
        "CT_Perspective",
    ):
        return "chart_config"
    if xtype in ("CT_Floor", "CT_SideWall", "CT_BackWall"):
        return "chart_config"
    if xtype in (
        "CT_DispBlanksAs",
        "CT_Grouping",
        "CT_BarGrouping",
        "CT_BarDir",
        "CT_Shape",
        "CT_Orientation",
        "CT_OfPieType",
        "CT_SplitType",
        "CT_BubbleScale",
        "CT_SizeRepresents",
        "CT_ScatterStyle",
        "CT_RadarStyle",
    ):
        return "chart_config"
    if xtype.startswith("CT_GapAmount") or xtype.startswith("CT_Overlap"):
        return "chart_config"
    if xtype.startswith("CT_UpDownBar"):
        return "chart_config"
    if xtype in ("CT_DataTable",):
        return "chart_config"

    if xtype in (
        "CT_PrintSettings",
        "CT_HeaderFooter",
        "CT_PageMargins",
        "CT_PageSetup",
        "CT_ExternalData",
    ):
        return "print_external"

    if xtype in ("CT_Protection", "CT_PivotFmt", "CT_PivotFmts", "CT_PivotSource"):
        return "protection_pivot"

    if xtype.startswith("CT_ExtensionList") or xtype.endswith("ExtensionList"):
        return "extensions"
    if "Extension" in cls and xtype != "CT_Extension":
        return "extensions"

    if xtype in (
        "CT_ChartSpace",
        "CT_Chart",
        "CT_PlotArea",
        "CT_PlotAreaRegion",
        "CT_Surface",
    ):
        return "structure"

    if xtype.startswith("CT_Num") and xtype not in (
        "CT_NumFmt",
        "CT_NumRef",
        "CT_NumData",
        "CT_NumDataSource",
        "CT_NumPoint",
    ):
        return "data_refs"

    if xtype.startswith("CT_Str") and xtype not in ("CT_StrRef", "CT_StrData"):
        return "data_refs"

    # Fallback
    return "other"


def main():
    print("Fetching schemas...", file=sys.stderr)

    all_schemas = {}
    all_types_by_label = {}

    for label, filename in SCHEMAS:
        try:
            schema = fetch_schema(filename)
            all_schemas[label] = (filename, schema)
            all_types_by_label[label] = schema.get("Types", [])
        except Exception as e:
            print(f"  WARNING: Failed to fetch {filename}: {e}", file=sys.stderr)
            all_schemas[label] = (filename, {"Types": []})
            all_types_by_label[label] = []

    # Collect a: refs from chart/drawing types
    a_refs = collect_a_refs(all_types_by_label)

    # Collect all enums across all schemas
    all_enums = set()
    for label, types_list in all_types_by_label.items():
        all_enums.update(extract_enum_refs(types_list))

    # Categorize c: types
    c_filename = all_schemas["c"][0]
    c_types = all_types_by_label.get("c", [])
    categories = defaultdict(list)
    for idx, t in enumerate(c_types):
        cat = categorize_chart_type(t)
        categories[cat].append((t, idx, c_filename))

    # Build output
    lines = []
    lines.append("# Chart Support — OpenXML Type Checklist")
    lines.append("")
    lines.append(
        "> Auto-generated from [Open-XML-SDK `data/schemas/`]"
        "(https://github.com/dotnet/Open-XML-SDK/tree/main/data/schemas)."
    )
    lines.append("> Each item references `schema_file#Types[index]` for traceability.")
    lines.append(">")
    lines.append("> Regenerate: `python3 tools/gen-chart-support.py`")
    lines.append("")

    # Section 1: Package structure (hand-written, these aren't in schemas)
    lines.append("## 1. Package Structure")
    lines.append("")
    lines.append("Relationships, content types, and part paths needed to embed")
    lines.append("charts in an XLSX file. These are not in the XSD schemas —")
    lines.append("they come from the SDK's part class definitions.")
    lines.append("")
    lines.append("- [x] `xl/drawings/drawingN.xml` \u2014 DrawingsPart")
    lines.append("- [x] `xl/charts/chartN.xml` \u2014 ChartPart")
    lines.append("- [ ] `xl/charts/styleN.xml` \u2014 ChartStylePart (Office 2013+)")
    lines.append("- [ ] `xl/charts/colorsN.xml` \u2014 ChartColorStylePart (Office 2013+)")
    lines.append("- [ ] `xl/chartsheets/sheetN.xml` \u2014 ChartsheetPart")
    lines.append("- [x] Worksheet \u2192 Drawing relationship (`RT_DRAWING`)")
    lines.append("- [x] Drawing \u2192 Chart relationship (`RT_CHART`)")
    lines.append("- [ ] Drawing \u2192 ChartEx relationship (`RT_CHART_EX`, Office 2016+)")
    lines.append("- [ ] Chart \u2192 ChartStyle relationship")
    lines.append("- [ ] Chart \u2192 ChartColorStyle relationship")
    lines.append("- [x] `[Content_Types].xml` override for DrawingsPart")
    lines.append("- [x] `[Content_Types].xml` override for ChartPart")
    lines.append("- [ ] `[Content_Types].xml` override for ChartStylePart")
    lines.append("- [ ] `[Content_Types].xml` override for ChartColorStylePart")
    lines.append('- [x] `<drawing r:id="..."/>` element in worksheet XML')
    lines.append("")

    # Section 2: SpreadsheetDrawing (xdr:)
    xdr_filename = all_schemas["xdr"][0]
    xdr_types = all_types_by_label.get("xdr", [])
    lines.append("## 2. SpreadsheetDrawing (`xdr:` namespace)")
    lines.append("")
    lines.append(f"Source: `{xdr_filename}`")
    lines.append("")
    for idx, t in enumerate(xdr_types):
        lines.append(format_item(t, idx, xdr_filename))
    lines.append("")

    # Section 3: Chart root & structure
    lines.append("## 3. Chart Root & Structure (`c:` namespace)")
    lines.append("")
    lines.append(f"Source: `{c_filename}`")
    lines.append("")
    for t, idx, fname in categories.get("structure", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 4: Chart types
    lines.append("## 4. Chart Types (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("chart_types", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 5: Series types
    lines.append("## 5. Series Types (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("series", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 6: Data references
    lines.append("## 6. Data References (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("data_refs", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 7: Axes
    lines.append("## 7. Axes (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("axes", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 8: Legend & title
    lines.append("## 8. Legend & Title (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("legend_title", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 9: Data labels & data points
    lines.append("## 9. Data Labels & Data Points (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("data_labels", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 10: Trendlines & error bars
    lines.append("## 10. Trendlines & Error Bars (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("trendlines", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 11: Formatting
    lines.append("## 11. Formatting (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("formatting", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 12: Chart configuration
    lines.append("## 12. Chart Configuration (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("chart_config", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 13: Boolean properties
    lines.append("## 13. Boolean Properties (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("booleans", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 14: Unsigned integer properties
    lines.append("## 14. Unsigned Integer Properties (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("unsigned_ints", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 15: Print settings & external data
    lines.append("## 15. Print Settings & External Data (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("print_external", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 16: Protection & pivot
    lines.append("## 16. Protection & Pivot (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("protection_pivot", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 17: Extension lists
    lines.append("## 17. Extension Lists (`c:` namespace)")
    lines.append("")
    for t, idx, fname in categories.get("extensions", []):
        lines.append(format_item(t, idx, fname))
    lines.append("")

    # Section 18: Other/uncategorized c: types
    if categories.get("other"):
        lines.append("## 18. Other (`c:` namespace)")
        lines.append("")
        for t, idx, fname in categories.get("other", []):
            lines.append(format_item(t, idx, fname))
        lines.append("")

    # Section 19: DrawingML subset (a: types referenced by charts)
    a_filename = all_schemas["a"][0]
    a_types = all_types_by_label.get("a", [])
    lines.append("## 19. DrawingML Subset (`a:` namespace)")
    lines.append("")
    lines.append("Only `a:` types directly referenced by chart or drawing elements.")
    lines.append(f"Source: `{a_filename}`")
    lines.append("")
    a_included = 0
    for idx, t in enumerate(a_types):
        name = t.get("Name", "")
        xtype_full = name.split("/")[0]
        if xtype_full in a_refs:
            lines.append(format_item(t, idx, a_filename))
            a_included += 1
    if a_included == 0:
        lines.append("_(no direct references found — check child element scan)_")
    lines.append("")

    # Section 20+: Extension schemas
    ext_section = 20
    for label in ("c14", "c15", "c16", "cx", "cs", "c16r3"):
        if label not in all_schemas:
            continue
        filename, schema = all_schemas[label]
        types_list = schema.get("Types", [])
        if not types_list:
            continue
        ns = schema.get("TargetNamespace", "")
        lines.append(f"## {ext_section}. Extensions: `{label}:` namespace")
        lines.append("")
        lines.append(f"Namespace: `{ns}`")
        lines.append(f"Source: `{filename}`")
        lines.append("")
        for idx, t in enumerate(types_list):
            lines.append(format_item(t, idx, filename))
        lines.append("")
        ext_section += 1

    # Section: Enumerations
    lines.append(f"## {ext_section}. Enumerations")
    lines.append("")
    lines.append("Enum types referenced in chart/drawing attribute definitions.")
    lines.append(
        "Values are defined in the SDK's generated code, not in the schema JSONs."
    )
    lines.append("")
    for enum_name in sorted(all_enums):
        lines.append(f"- [ ] `{enum_name}`")
    lines.append("")

    # Summary
    total_checked = sum(1 for l in lines if l.startswith("- [x]"))
    total_unchecked = sum(1 for l in lines if l.startswith("- [ ]"))
    total_checkboxes = total_checked + total_unchecked
    total_abstract = sum(1 for l in lines if "_abstract base:" in l)
    lines.insert(
        7,
        f"**Total items: {total_checkboxes} checkboxes "
        f"({total_checked} implemented, {total_unchecked} remaining), "
        f"{total_abstract} abstract bases**",
    )
    lines.insert(8, "")

    output = "\n".join(lines) + "\n"

    outpath = "docs/CHART_SUPPORT.md"
    with open(outpath, "w") as f:
        f.write(output)

    print(
        f"\nWrote {outpath} ({total_checked} implemented, "
        f"{total_unchecked} remaining, {total_abstract} abstract bases, "
        f"{len(lines)} lines)",
        file=sys.stderr,
    )


if __name__ == "__main__":
    main()
