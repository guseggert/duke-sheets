#!/usr/bin/env python3
"""Check ChartEx implementation coverage across spec, Rust model, and bindings.

Usage: python3 tools/check-chartex-coverage.py

Reads:
  - /tmp/chartex-schema.json (Open-XML-SDK schema)
  - crates/duke-sheets-chart/src/chart_ex.rs (Rust model)
  - bindings/nodejs/src/types.rs (Node.js binding)
  - bindings/python/src/types.rs (Python binding)
  - bindings/wasm/src/types.rs (WASM binding)

Prints a coverage report showing gaps at each layer.
"""

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCHEMA_PATH = Path("/tmp/chartex-schema.json")
MODEL_PATH = ROOT / "crates/duke-sheets-chart/src/chart_ex.rs"
BINDINGS = {
    "nodejs": ROOT / "bindings/nodejs/src/types.rs",
    "python": ROOT / "bindings/python/src/types.rs",
    "wasm": ROOT / "bindings/wasm/src/types.rs",
}


def parse_spec(path):
    """Parse the Open-XML-SDK schema JSON into a structured dict."""
    with open(path) as f:
        schema = json.load(f)

    types = {}
    for t in schema["Types"]:
        name = t["Name"]
        ct_name = name.split("/")[0]  # e.g. "cx:CT_Series"
        element = name.split("/")[-1] if "/" in name else ""
        class_name = t.get("ClassName", "")

        # Skip abstract types, leaf text types, DrawingML reuse types
        if t.get("IsAbstract"):
            continue
        if t.get("IsLeafText") and not element.startswith("cx:"):
            continue
        if ct_name.startswith("a:"):
            continue
        if ct_name.startswith("c:"):
            continue
        if ct_name.startswith("xsd:"):
            continue

        attrs = []
        for a in t.get("Attributes", []):
            qname = a["QName"]
            attr_name = qname.lstrip(":")
            if attr_name.startswith("cx:"):
                attr_name = attr_name[3:]
            attrs.append(attr_name)

        children = []
        for c in t.get("Children", []):
            child_el = c["Name"].split("/")[-1]
            children.append(child_el)

        types[class_name] = {
            "ct_name": ct_name,
            "element": element,
            "attrs": attrs,
            "children": children,
            "is_leaf": t.get("IsLeafElement", False) or t.get("IsLeafText", False),
        }

    enums = {}
    for e in schema.get("Enums", []):
        values = [f["Value"] for f in e.get("Facets", [])]
        enums[e["Name"]] = values

    return types, enums


def parse_rust_model(path):
    """Parse Rust struct/enum definitions from chart_ex.rs."""
    with open(path) as f:
        content = f.read()

    types = {}
    current = None
    current_kind = None

    for line in content.split("\n"):
        m = re.match(r"pub (struct|enum) (\w+)", line)
        if m:
            current = m.group(2)
            current_kind = m.group(1)
            types[current] = {"kind": current_kind, "fields": [], "raw_fields": []}
            continue

        if current and current_kind == "struct":
            m2 = re.match(r"\s+pub (\w+):\s*(.*?)(?:,\s*)?$", line)
            if m2:
                field_name = m2.group(1)
                field_type = m2.group(2).strip().rstrip(",")
                types[current]["fields"].append(field_name)
                types[current]["raw_fields"].append((field_name, field_type))

        if current and current_kind == "enum":
            m3 = re.match(r"\s+(\w+)[\s,({]", line)
            if m3:
                variant = m3.group(1)
                if variant not in (
                    "pub",
                    "use",
                    "fn",
                    "impl",
                    "let",
                    "if",
                    "for",
                    "match",
                    "Self",
                    "Some",
                    "None",
                    "Ok",
                    "Err",
                ):
                    types[current]["fields"].append(variant)

    return types


def parse_binding(path, prefix):
    """Parse binding types.rs for ChartEx-related structs."""
    if not path.exists():
        return {}

    with open(path) as f:
        content = f.read()

    types = {}
    current = None

    for line in content.split("\n"):
        m = re.match(rf"pub struct {prefix}(\w*ChartEx\w*)", line)
        if m:
            current = m.group(1)
            types[current] = []
            continue

        if current:
            m2 = re.match(r"\s+pub (\w+):", line)
            if m2:
                types[current].append(m2.group(1))
            elif line.strip() == "}" or (
                line.strip().startswith("pub struct") and "ChartEx" not in line
            ):
                current = None

    # Detect ChartExLayout exposed via string conversion function
    if re.search(r"fn chart_ex_layout_to_string", content):
        types["ChartExLayout"] = ["__string_enum__"]

    return types

# Mapping from spec ClassName to our Rust struct name
SPEC_TO_RUST = {
    "ChartSpace": "ChartEx",
    "Chart": None,  # folded into ChartEx
    "ChartTitle": "ChartExTitle",
    "ChartData": None,  # container, children are ChartExData
    "Data": "ChartExData",
    "NumericDimension": "ChartExDimension",  # enum variant
    "StringDimension": "ChartExDimension",  # enum variant
    "NumericLevel": "ChartExNumericLevel",
    "StringLevel": "ChartExStringLevel",
    "NumericValue": None,  # leaf, inside level.points
    "ChartStringValue": None,  # leaf, inside level.points
    "PlotArea": "ChartExPlotArea",
    "PlotAreaRegion": None,  # folded into PlotArea
    "PlotSurface": None,  # field on PlotArea
    "Series": "ChartExSeries",
    "SeriesLayoutProperties": "ChartExLayoutPr",
    "Axis": "ChartExAxis",
    "AxisTitle": "ChartExAxisTitle",
    "AxisUnits": "ChartExAxisUnits",
    "AxisUnitsLabel": "ChartExAxisUnitsLabel",
    "Legend": "ChartExLegend",
    "DataLabels": "ChartExDataLabels",
    "DataLabel": "ChartExDataLabel",
    "DataLabelHidden": None,  # stored as hidden_labels Vec<u32> on DataLabels
    "DataPoint": "ChartExDataPoint",
    "DataLabelVisibilities": None,  # inlined as fields on DataLabels/DataLabel
    "Text": "ChartExText",
    "TextData": "ChartExTextData",
    "RichTextBody": None,  # stored as raw bytes
    "Formula": None,  # leaf text
    "NfFormula": None,  # leaf text
    "VXsdstring": None,  # leaf text
    "NumberFormat": None,  # reuse existing NumberFormat
    "ShapeProperties": None,  # reuse existing ChartShapeProperties
    "TxPrTextBody": None,  # reuse existing TextProperties
    "Extension2": None,  # raw passthrough
    "ExtensionList": None,  # raw passthrough
    "ExternalData": "ChartExExternalData",
    "RelId": None,  # drawing-level, not in model
    "Geography": "ChartExGeography",
    "Subtotals": None,  # inlined as Vec<u32> on LayoutPr
    "Binning": "ChartExBinning",
    "Statistics": "ChartExStatistics",
    "Aggregation": None,  # bool field on LayoutPr
    "ParentLabelLayout": None,  # string field on LayoutPr
    "RegionLabelLayout": None,  # string field on LayoutPr
    "SeriesElementVisibilities": "ChartExSeriesVisibility",
    "ValueColors": "ChartExValueColors",
    "ValueColorPositions": "ChartExValueColorPositions",
    "MinValueColorEndPosition": None,  # field on ValueColorPositions
    "MaxValueColorEndPosition": None,  # field on ValueColorPositions
    "ValueColorMiddlePosition": None,  # field on ValueColorPositions
    "ExtremeValueColorPosition": None,  # enum variant
    "NumberColorPosition": None,  # enum variant
    "PercentageColorPosition": None,  # enum variant
    "FormatOverride": "ChartExFormatOverride",
    "FormatOverrides": None,  # container
    "PrintSettings": "ChartExPrintSettings",
    "HeaderFooter": "ChartExHeaderFooter",
    "PageMargins": "ChartExPageMargins",
    "PageSetup": "ChartExPageSetup",
    "Offset": "ChartExOffset",
    "CategoryAxisScaling": None,  # enum variant on ChartExScaling
    "ValueAxisScaling": None,  # enum variant on ChartExScaling
    "MajorGridlinesGridlines": None,  # field on Axis
    "MinorGridlinesGridlines": None,  # field on Axis
    "MajorTickMarksTickMarks": None,  # field on Axis
    "MinorTickMarksTickMarks": None,  # field on Axis
    "TickLabels": None,  # bool field on Axis
    "GeoCache": None,  # raw bytes inside Geography
    "Clear": None,  # inside GeoCache, raw
    "Copyrights": None,  # inside GeoCache, raw
    "DataId": None,  # leaf, stored as u32 on Series
    "Xsdbase64Binary": None,  # inside GeoCache
    "Xsddouble": None,  # leaf for binSize
    "BinCountXsdunsignedInt": None,  # leaf for binCount
    "UnsignedIntegerType": None,  # cx:idx, leaf
    "EntityType": None,  # inside geo types
    "SeparatorXsdstring": None,  # leaf
    "OddHeaderXsdstring": None,  # leaf
    "OddFooterXsdstring": None,  # leaf
    "EvenHeaderXsdstring": None,  # leaf
    "EvenFooterXsdstring": None,  # leaf
    "FirstHeaderXsdstring": None,  # leaf
    "FirstFooterXsdstring": None,  # leaf
    "CopyrightXsdstring": None,  # leaf
    "ColorMappingType": None,  # raw bytes
    # Geo types - all raw passthrough inside geoCache
    "GeoData": None,
    "GeoDataEntityQuery": None,
    "GeoDataEntityQueryResult": None,
    "GeoDataEntityQueryResults": None,
    "GeoDataPointQuery": None,
    "GeoDataPointToEntityQuery": None,
    "GeoDataPointToEntityQueryResult": None,
    "GeoDataPointToEntityQueryResults": None,
    "GeoEntity": None,
    "GeoHierarchyEntity": None,
    "GeoLocation": None,
    "GeoLocationQuery": None,
    "GeoLocationQueryResult": None,
    "GeoLocationQueryResults": None,
    "GeoLocations": None,
    "GeoChildEntities": None,
    "GeoChildEntitiesQuery": None,
    "GeoChildEntitiesQueryResult": None,
    "GeoChildEntitiesQueryResults": None,
    "GeoChildTypes": None,
    "GeoParentEntitiesQuery": None,
    "GeoParentEntitiesQueryResult": None,
    "GeoParentEntitiesQueryResults": None,
    "GeoParentEntity": None,
    "GeoPolygon": None,
    "GeoPolygons": None,
    "Address": None,
    "MinColorSolidColorFillProperties": None,  # DrawingML
    "MidColorSolidColorFillProperties": None,  # DrawingML
    "MaxColorSolidColorFillProperties": None,  # DrawingML
    "OpenXmlSolidColorFillPropertiesElement": None,  # abstract
    "TextBodyType": None,  # abstract
    "OpenXmlFormulaElement": None,  # abstract
    "Openxmlsdk_49BECFFA_3B03_4D13_8272_D6CCB22579E3XsdunsignedInt": None,  # SDK internal
    "AxisId": None,  # leaf
}

# Mapping from Rust model field to expected spec attribute/child
# (for the types where we DO model them)
RUST_FIELD_TO_SPEC = {
    "ChartEx": {
        "title": "cx:title",
        "data": "cx:chartData",  # indirect
        "plot_area": "cx:plotArea",  # indirect via cx:chart
        "legend": "cx:legend",  # indirect via cx:chart
        "anchor": None,  # drawing-level, not in chartSpace
        "shape_properties": "cx:spPr",
        "text_properties": "cx:txPr",
        "color_map_override": "cx:clrMapOvr",
        "format_overrides": "cx:fmtOvrs",
        "print_settings": "cx:printSettings",
        "raw_chart_style": None,  # separate part
        "raw_chart_color_style": None,  # separate part
        "raw_extensions": "cx:extLst",
        "raw_mc_fallback": None,  # drawing-level
    },
}


def main():
    if not SCHEMA_PATH.exists():
        print(f"Schema not found at {SCHEMA_PATH}")
        print(
            "Run: curl -sL 'https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/schemas/schemas_microsoft_com_office_drawing_2014_chartex.json' -o /tmp/chartex-schema.json"
        )
        sys.exit(1)

    spec_types, spec_enums = parse_spec(SCHEMA_PATH)
    rust_types = parse_rust_model(MODEL_PATH)

    print("=" * 70)
    print("ChartEx Coverage Report")
    print("=" * 70)

    # Layer 1: Spec → Rust model
    print("\n## Layer 1: Spec Types → Rust Model")
    print()
    unmapped = []
    mapped_with_rust = []
    explicitly_skipped = []

    for class_name in sorted(spec_types.keys()):
        if class_name in SPEC_TO_RUST:
            rust_name = SPEC_TO_RUST[class_name]
            if rust_name is None:
                explicitly_skipped.append(class_name)
            elif rust_name in rust_types:
                mapped_with_rust.append((class_name, rust_name))
            else:
                unmapped.append(
                    (class_name, f"mapped to {rust_name} but NOT FOUND in Rust")
                )
        else:
            unmapped.append((class_name, "NO MAPPING defined"))

    print(f"  Spec types:         {len(spec_types)}")
    print(f"  Mapped to Rust:     {len(mapped_with_rust)}")
    print(
        f"  Explicitly skipped: {len(explicitly_skipped)}"
    )
    print(f"  UNMAPPED:           {len(unmapped)}")
    if unmapped:
        print()
        for class_name, reason in unmapped:
            spec = spec_types[class_name]
            print(f"  ❌ {class_name} ({spec['element']}): {reason}")
            if spec["attrs"]:
                print(f"     attrs: {spec['attrs']}")
            if spec["children"]:
                print(f"     children: {spec['children']}")

    # Categorize skipped types
    leaf_text = []
    geo = []
    drawingml = []
    folded = []
    for cn in explicitly_skipped:
        cnl = cn.lower()
        if any(g in cnl for g in ('geo', 'address', 'polygon', 'copyright', 'entity')):
            geo.append(cn)
        elif any(d in cnl for d in ('solid', 'textbody', 'colormapping', 'openxml')):
            drawingml.append(cn)
        elif cn in ('Chart', 'ChartData', 'PlotAreaRegion', 'PlotSurface', 'FormatOverrides',
                    'Subtotals', 'Aggregation', 'ParentLabelLayout', 'RegionLabelLayout',
                    'DataLabelVisibilities', 'DataLabelHidden', 'DataId'):
            folded.append(cn)
        else:
            leaf_text.append(cn)

    print(f"\n  Skipped type breakdown:")
    print(f"    Leaf/text types (no struct needed, values on parent): {len(leaf_text)}")
    print(f"    Geo types (raw bytes in geoCache):                   {len(geo)}")
    print(f"    Folded into parent structs:                          {len(folded)}")
    print(f"    DrawingML reuse (existing types):                     {len(drawingml)}")

    # Layer 1b: Check field coverage for mapped types
    print(f"\n## Layer 1b: Field Coverage (Spec attrs/children → Rust fields)")
    print()
    total_spec_fields = 0
    total_rust_fields = 0
    field_gaps = []

    for class_name, rust_name in mapped_with_rust:
        spec = spec_types[class_name]
        rust = rust_types[rust_name]
        spec_fields = spec["attrs"] + [c.replace("cx:", "") for c in spec["children"]]
        rust_fields = rust["fields"]
        total_spec_fields += len(spec_fields)
        total_rust_fields += len(rust_fields)

        # Known name mappings from spec element/attr names to Rust field names
        FIELD_ALIASES = {
            "spPr": "shape_properties",
            "txPr": "text_properties",
            "numFmt": "number_format",
            "extLst": "extensions",
            "r:id": "rel_id",
            "cx:autoUpdate": "auto_update",
            "f": "formula",
            "nf": "nf_formula",
            "v": "value",
            "tx": "text",
            "lvl": "levels",
            "plotAreaRegion": "series",
            "axis": "axes",
            "dataPt": "data_points",
            "axisId": "axis_ids",
            "numDim": "dimensions",
            "strDim": "dimensions",
            "dataLabel": "overrides",
            "dataLabelHidden": "hidden_labels",
            "geoCache": "raw_geo_cache",
            "version": "version",
            "featureList": "feature_list",
            "fallbackImg": "fallback_img",
            "clrMapOvr": "color_map_override",
            "fmtOvrs": "format_overrides",
        }

        missing = []
        for sf in spec_fields:
            alias = FIELD_ALIASES.get(sf)
            if alias and alias in rust_fields:
                continue
            snake = re.sub(r"([A-Z])", r"_\1", sf).lower().lstrip("_")
            found = any(
                rf == snake or rf == sf.lower() or sf.lower() in rf or rf in sf.lower()
                for rf in rust_fields
            )
            if not found:
                missing.append(sf)
        if missing:
            field_gaps.append((class_name, rust_name, missing))

    if field_gaps:
        for class_name, rust_name, missing in field_gaps:
            print(f"  ⚠️  {class_name} → {rust_name}: missing fields: {missing}")
    else:
        print("  ✅ All mapped types have field coverage")
    print(f"\n  Spec fields total:  {total_spec_fields}")
    print(f"  Rust fields total:  {total_rust_fields}")

    # Layer 2: Rust model → Bindings
    print(f"\n## Layer 2: Rust Model → Bindings")
    print()

    # Find which Rust types should be in bindings
    # (public types that represent user-facing chart data)
    binding_types = [
        "ChartEx",
        "ChartExSeries",
        "ChartExData",
        "ChartExDimension",
        "ChartExAxis",
        "ChartExLegend",
        "ChartExDataLabels",
        "ChartExTitle",
        "ChartExLayoutPr",
        "ChartExLayout",
    ]

    # Fields to skip in binding coverage (raw bytes, internal, TextProperties)
    SKIP_FIELDS = (
        "raw_chart_style",
        "raw_chart_color_style",
        "raw_extensions",
        "raw_mc_fallback",
        "raw_geo_cache",
        "color_map_override",
        "extensions",
        "text_properties",
        "rich_text",
        "rich",
    )

    # Enum types exposed as flattened structs or string conversions -
    # skip field-level checks ("fields" are variant names, not struct fields)
    ENUM_TYPES = {"ChartExLayout", "ChartExDimension"}

    for binding_name, binding_path in sorted(BINDINGS.items()):
        prefix = {"nodejs": "Js", "python": "Py", "wasm": "Wasm"}[binding_name]
        binding = parse_binding(binding_path, prefix)

        print(f"  ### {binding_name} ({binding_path.name})")
        present = []
        missing_types = []
        for rt in binding_types:
            bt = rt  # binding type name without prefix
            if bt in binding:
                present.append(bt)
                # Skip field-level checks for enum types
                if rt in ENUM_TYPES:
                    continue
                # Check field coverage
                rust_fields = rust_types.get(rt, {}).get("fields", [])
                binding_fields = binding[bt]
                missing_fields = [
                    f
                    for f in rust_fields
                    if f not in binding_fields
                    and f not in SKIP_FIELDS
                    # Handle flattened fields: external_data → external_data_rel_id + external_data_auto_update
                    and not any(bf.startswith(f + "_") for bf in binding_fields)
                ]
                if missing_fields:
                    print(f"    ⚠️  {prefix}{bt}: missing fields: {missing_fields}")
            else:
                missing_types.append(rt)

        if missing_types:
            for mt in missing_types:
                print(f"    ❌ {prefix}{mt}: NOT EXPOSED")
        if not missing_types and not any(
            f
            for rt in present
            if rt not in ENUM_TYPES
            for f in rust_types.get(rt, {}).get("fields", [])
            if f not in binding.get(rt, [])
            and f not in SKIP_FIELDS
        ):
            print(f"    ✅ All expected types present")
        print(f"    Types exposed: {len(present)}/{len(binding_types)}")
        print()

    # Layer 3: Spec enums → Rust enums
    print(f"## Layer 3: Spec Enums")
    print()
    rust_enums = {k: v for k, v in rust_types.items() if v["kind"] == "enum"}
    print(f"  Spec enums: {len(spec_enums)}")
    print(f"  Rust enums: {len(rust_enums)}")
    print()
    enum_map = {
        "SeriesLayout": "ChartExLayout",
        "NumericDimensionType": "NumericDimType",
        "StringDimensionType": "StringDimType",
    }
    for spec_name, values in sorted(spec_enums.items()):
        rust_name = enum_map.get(spec_name)
        if rust_name and rust_name in rust_enums:
            rust_variants = [v.lower() for v in rust_enums[rust_name]["fields"]]
            missing = [
                v
                for v in values
                if v.lower() not in rust_variants and v not in ("unknown",)
            ]
            if missing:
                print(f"  ⚠️  {spec_name} → {rust_name}: missing variants: {missing}")
            else:
                print(
                    f"  ✅ {spec_name} → {rust_name}: all {len(values)} values covered"
                )
        elif rust_name:
            print(f"  ❌ {spec_name} → {rust_name}: NOT FOUND")
        else:
            # Enums stored as strings - check that the plan says so
            print(f"  ℹ️  {spec_name}: stored as String (not a Rust enum)")


if __name__ == "__main__":
    main()
