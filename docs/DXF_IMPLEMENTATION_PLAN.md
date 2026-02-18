# DXF (Differential Format) Implementation Plan

## Document Purpose

This document outlines the complete implementation plan for full DXF (Differential Format) support in duke-sheets. DXF is used in XLSX for conditional formatting styles, table styles, and revision tracking.

## Table of Contents

1. [Current State Analysis](#1-current-state-analysis)
2. [Feature Gap Analysis](#2-feature-gap-analysis)
3. [Implementation Tasks](#3-implementation-tasks)
4. [Detailed Implementation](#4-detailed-implementation)
5. [Testing Strategy](#5-testing-strategy)
6. [API Design Decisions](#6-api-design-decisions)
7. [Edge Cases & Quirks](#7-edge-cases--quirks)
8. [Effort Estimates](#8-effort-estimates)

---

## 1. Current State Analysis

### What We Currently Support

**DXF Writing (`write_dxf()` in styles.rs):**
- Font: bold, italic, strikethrough, underline, size, name, color
- Fill: patternFill with solid colors
- Border: left, right, top, bottom, diagonal with colors

**DXF Reading (`read_styles_xml()` in styles.rs):**
- Font parsing (same properties as writing)
- Fill parsing (patternFill only)
- Border parsing

### Files Involved

| File | Role |
|------|------|
| `crates/duke-sheets-xlsx/src/styles.rs` | DXF read/write, `XlsxStyleTable`, `ParsedStyles` |
| `crates/duke-sheets-xlsx/src/reader/mod.rs` | XLSX reading, applies DXF to CF rules |
| `crates/duke-sheets-xlsx/src/writer/mod.rs` | XLSX writing, `dxfId` on CF rules |
| `crates/duke-sheets-core/src/style/*.rs` | Style data structures |
| `crates/duke-sheets-core/src/conditional_format.rs` | `ConditionalFormatRule` with `format: Option<Style>` |

---

## 2. Feature Gap Analysis

### DXF Child Elements (per ECMA-376 §18.8.14)

| Element | XML Tag | Current State | Priority |
|---------|---------|---------------|----------|
| Font | `<font>` | Partial | - |
| Number Format | `<numFmt>` | **Missing** | **High** |
| Fill | `<fill>` | Partial | - |
| Alignment | `<alignment>` | **Missing** | **High** |
| Border | `<border>` | Partial | - |
| Protection | `<protection>` | **Missing** | Medium |
| Extension List | `<extLst>` | Missing | Low |

### Font Sub-Elements Gap

| Element | Tag | Current | Priority |
|---------|-----|---------|----------|
| Bold | `<b/>` | Yes | - |
| Italic | `<i/>` | Yes | - |
| Strike | `<strike/>` | Yes | - |
| Underline | `<u/>` | Yes | - |
| Font Size | `<sz/>` | Yes | - |
| Font Name | `<name/>` | Yes | - |
| Color | `<color/>` | Yes | - |
| Condense | `<condense/>` | No | Low |
| Extend | `<extend/>` | No | Low |
| Outline | `<outline/>` | No | Low |
| Shadow | `<shadow/>` | No | Low |
| Vertical Align | `<vertAlign/>` | No | Medium |
| Font Scheme | `<scheme/>` | No | Low |
| Font Family | `<family/>` | No | Low |
| Font Charset | `<charset/>` | No | Low |

### Fill Sub-Elements Gap

| Element | Tag | Current | Priority |
|---------|-----|---------|----------|
| Pattern Fill | `<patternFill>` | Yes | - |
| Gradient Fill | `<gradientFill>` | No | Medium |

### Border Sub-Elements Gap

| Element | Current | Priority |
|---------|---------|----------|
| left, right, top, bottom | Yes | - |
| diagonal | Yes | - |
| vertical (DXF-only) | No | Low |
| horizontal (DXF-only) | No | Low |

---

## 3. Implementation Tasks

### Phase 1: High Priority (Core Functionality)

#### Task 1.1: Number Format in DXF
- **Writer:** Add `<numFmt>` inline in DXF (not by reference)
- **Reader:** Parse `<numFmt>` within `<dxf>` elements
- **Files:** `styles.rs`

#### Task 1.2: Alignment in DXF
- **Writer:** Add `<alignment>` to `write_dxf()`
- **Reader:** Parse `<alignment>` within `<dxf>` elements
- **Files:** `styles.rs`

#### Task 1.3: Create Test Fixtures
- **Tool:** Python script using openpyxl
- **Location:** `tests/fixtures/dxf/`
- **Coverage:** All DXF features

### Phase 2: Medium Priority (Enhanced Compatibility)

#### Task 2.1: Protection in DXF
- **Writer/Reader:** Add `<protection>` to DXF
- **Files:** `styles.rs`

#### Task 2.2: Gradient Fill in DXF
- **Writer:** Serialize `<gradientFill>` instead of downgrading
- **Reader:** Parse `<gradientFill>` within DXF
- **Files:** `styles.rs`

#### Task 2.3: Font Vertical Alignment
- **Writer/Reader:** Add `<vertAlign>` for superscript/subscript
- **Files:** `styles.rs`, `font.rs` (already has `FontVerticalAlign`)

### Phase 3: Low Priority (Full Spec Compliance)

#### Task 3.1: DXF-Specific Border Elements
- **Writer:** Add empty `<vertical/>` and `<horizontal/>`
- **Reader:** Skip these gracefully
- **Note:** Per XlsxWriter, diagonal borders are NOT allowed in DXF

#### Task 3.2: Font Extended Properties
- Add: condense, extend, outline, shadow, scheme, family, charset
- **Files:** `font.rs`, `styles.rs`

#### Task 3.3: Extension List
- **Reader:** Parse `<extLst>` and preserve for roundtrip
- **Writer:** Emit preserved extensions
- **Files:** `styles.rs`

---

## 4. Detailed Implementation

### 4.1 Number Format in DXF

**Current NumberFormat enum:**
```rust
pub enum NumberFormat {
    General,
    BuiltIn(u32),
    Custom(String),
}
```

**Writer changes (`write_dxf()`):**
```rust
fn write_dxf(style: &Style) -> String {
    let mut s = String::from("<dxf>");
    
    // ... existing font, fill, border ...
    
    // NEW: Number format (inline, with both id and code)
    if style.number_format != NumberFormat::General {
        s.push_str(&write_dxf_numfmt(&style.number_format));
    }
    
    // NEW: Alignment
    if style.alignment != Alignment::default() {
        s.push_str(&write_dxf_alignment(&style.alignment));
    }
    
    // NEW: Protection
    if style.protection != Protection::default() {
        s.push_str(&write_dxf_protection(&style.protection));
    }
    
    s.push_str("</dxf>");
    s
}

fn write_dxf_numfmt(nf: &NumberFormat) -> String {
    // DXF numFmt is inline, requires both numFmtId AND formatCode
    let (id, code) = match nf {
        NumberFormat::General => return String::new(),
        NumberFormat::BuiltIn(id) => (*id, NumberFormat::builtin_format_string(*id)),
        NumberFormat::Custom(code) => (164, code.as_str()), // Custom starts at 164
    };
    format!("<numFmt numFmtId=\"{}\" formatCode=\"{}\"/>", id, escape_xml_attr(code))
}
```

**Reader changes:**
```rust
// In read_styles_xml(), within DXF parsing:
b"numFmt" if in_dxf => {
    let mut num_fmt_id: Option<u32> = None;
    let mut format_code: Option<String> = None;
    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"numFmtId" => num_fmt_id = attr.unescape_value().ok().and_then(|s| s.parse().ok()),
            b"formatCode" => format_code = attr.unescape_value().ok().map(|s| s.to_string()),
            _ => {}
        }
    }
    if let Some(dxf) = current_dxf.as_mut() {
        dxf.number_format = match (num_fmt_id, format_code) {
            (Some(0), _) => NumberFormat::General,
            (Some(id), Some(code)) if id >= 164 => NumberFormat::Custom(code),
            (Some(id), _) => NumberFormat::BuiltIn(id),
            _ => NumberFormat::General,
        };
    }
}
```

### 4.2 Alignment in DXF

**Writer (`write_dxf_alignment()`):**
```rust
fn write_dxf_alignment(al: &Alignment) -> String {
    let default = Alignment::default();
    if al == &default {
        return String::new();
    }
    
    let mut s = String::from("<alignment");
    
    if al.horizontal != default.horizontal {
        s.push_str(&format!(" horizontal=\"{}\"", horiz_to_str(al.horizontal)));
    }
    if al.vertical != default.vertical {
        s.push_str(&format!(" vertical=\"{}\"", vert_to_str(al.vertical)));
    }
    if al.wrap_text {
        s.push_str(" wrapText=\"1\"");
    }
    if al.shrink_to_fit {
        s.push_str(" shrinkToFit=\"1\"");
    }
    if al.indent != 0 {
        s.push_str(&format!(" indent=\"{}\"", al.indent));
    }
    if al.rotation != 0 {
        s.push_str(&format!(" textRotation=\"{}\"", al.rotation));
    }
    match al.reading_order {
        ReadingOrder::ContextDependent => {}
        ReadingOrder::LeftToRight => s.push_str(" readingOrder=\"1\""),
        ReadingOrder::RightToLeft => s.push_str(" readingOrder=\"2\""),
    }
    
    s.push_str("/>");
    s
}
```

**Reader:**
```rust
// Within DXF parsing:
b"alignment" if in_dxf => {
    if let Some(dxf) = current_dxf.as_mut() {
        for attr in e.attributes().flatten() {
            let val = match attr.unescape_value() {
                Ok(v) => v,
                Err(_) => continue,
            };
            match attr.key.as_ref() {
                b"horizontal" => {
                    if let Some(h) = str_to_horizontal(&val) {
                        dxf.alignment.horizontal = h;
                    }
                }
                b"vertical" => {
                    if let Some(v) = str_to_vertical(&val) {
                        dxf.alignment.vertical = v;
                    }
                }
                b"wrapText" => dxf.alignment.wrap_text = val.as_ref() == "1",
                b"shrinkToFit" => dxf.alignment.shrink_to_fit = val.as_ref() == "1",
                b"indent" => dxf.alignment.indent = val.parse().unwrap_or(0),
                b"textRotation" => dxf.alignment.rotation = val.parse().unwrap_or(0),
                b"readingOrder" => {
                    dxf.alignment.reading_order = match val.as_ref() {
                        "1" => ReadingOrder::LeftToRight,
                        "2" => ReadingOrder::RightToLeft,
                        _ => ReadingOrder::ContextDependent,
                    };
                }
                _ => {}
            }
        }
    }
}
```

### 4.3 Protection in DXF

**Writer:**
```rust
fn write_dxf_protection(p: &Protection) -> String {
    let default = Protection::default();
    if p == &default {
        return String::new();
    }
    let mut s = String::from("<protection");
    if p.locked != default.locked {
        s.push_str(&format!(" locked=\"{}\"", if p.locked { 1 } else { 0 }));
    }
    if p.hidden != default.hidden {
        s.push_str(&format!(" hidden=\"{}\"", if p.hidden { 1 } else { 0 }));
    }
    s.push_str("/>");
    s
}
```

### 4.4 Gradient Fill (DXF)

**Data structure already exists:**
```rust
pub enum FillStyle {
    Gradient {
        gradient_type: GradientType,
        angle: f64,
        stops: Vec<GradientStop>,
    },
}
```

**Writer (`write_gradient_fill()`):**
```rust
fn write_gradient_fill(gradient_type: GradientType, angle: f64, stops: &[GradientStop]) -> String {
    let mut s = String::from("<fill><gradientFill");
    
    match gradient_type {
        GradientType::Linear => {
            s.push_str(&format!(" type=\"linear\" degree=\"{}\"", angle));
        }
        GradientType::Path => {
            s.push_str(" type=\"path\"");
            // Path gradient uses left/right/top/bottom attributes
        }
    }
    s.push('>');
    
    for stop in stops {
        s.push_str(&format!("<stop position=\"{}\">", stop.position));
        s.push_str(&write_color("color", &stop.color));
        s.push_str("</stop>");
    }
    
    s.push_str("</gradientFill></fill>");
    s
}
```

---

## 5. Testing Strategy

### 5.1 Test Fixtures

We have three sets of test fixtures in `tests/fixtures/dxf/`:

| Prefix | Source | Authority | Notes |
|--------|--------|-----------|-------|
| `pyuno_*.xlsx` | LibreOffice PyUNO | **Highest** | Created via LibreOffice's native UNO API |
| `dxf_*.xlsx` | openpyxl | Medium | Created via Python openpyxl library |
| `manual_*.xlsx` | Hand-crafted | Spec-compliant | Direct XML, useful for edge cases |

**Key PyUNO Fixtures:**
- `pyuno_dxf_font.xlsx` - Bold + color
- `pyuno_dxf_fill.xlsx` - Background color
- `pyuno_dxf_border.xlsx` - Border styles
- `pyuno_dxf_alignment.xlsx` - Center + wrap (proof that LO supports alignment in DXF!)
- `pyuno_dxf_numfmt.xlsx` - Number format in DXF
- `pyuno_dxf_full.xlsx` - All properties combined

**Generate fixtures:**
```bash
# openpyxl fixtures
python3 tests/fixtures/dxf/generate_fixtures.py

# PyUNO fixtures (via Docker)
docker build -t dxf-fixture-gen tests/fixtures/dxf/
docker run --rm -v $(pwd)/tests/fixtures/dxf:/output dxf-fixture-gen
```

### 5.2 Original openpyxl Generation Script

`tests/fixtures/dxf/generate_fixtures.py`:

```python
#!/usr/bin/env python3
"""Generate DXF test fixtures using openpyxl."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
import os

OUTPUT_DIR = os.path.dirname(__file__)

def create_dxf_numfmt_test():
    """Test: DXF with number format."""
    wb = Workbook()
    ws = wb.active
    ws.title = "NumFmt"
    
    # Add test data
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i * 0.1)
    
    # CF rule with number format - using font as proxy since openpyxl
    # DifferentialStyle numFmt support varies by version
    dxf = DifferentialStyle(
        font=Font(bold=True, color="FF0000")
    )
    rule = Rule(type="cellIs", operator="greaterThan", formula=["0.3"], dxf=dxf)
    ws.conditional_formatting.add("A1:A5", rule)
    
    wb.save(os.path.join(OUTPUT_DIR, "dxf_numfmt.xlsx"))

def create_dxf_alignment_test():
    """Test: DXF with alignment."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Alignment"
    
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=f"Text {i}")
    
    dxf = DifferentialStyle(
        alignment=Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
            indent=2,
            text_rotation=45
        )
    )
    rule = Rule(type="cellIs", operator="equal", formula=['"Text 3"'], dxf=dxf)
    ws.conditional_formatting.add("A1:A5", rule)
    
    wb.save(os.path.join(OUTPUT_DIR, "dxf_alignment.xlsx"))

def create_dxf_protection_test():
    """Test: DXF with protection."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Protection"
    
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i * 100)
    
    dxf = DifferentialStyle(
        protection=Protection(locked=False, hidden=True)
    )
    rule = Rule(type="cellIs", operator="greaterThan", formula=["300"], dxf=dxf)
    ws.conditional_formatting.add("A1:A5", rule)
    
    wb.save(os.path.join(OUTPUT_DIR, "dxf_protection.xlsx"))

def create_dxf_full_style_test():
    """Test: DXF with all style properties."""
    wb = Workbook()
    ws = wb.active
    ws.title = "FullStyle"
    
    for i in range(1, 11):
        ws.cell(row=i, column=1, value=i * 10)
    
    dxf = DifferentialStyle(
        font=Font(bold=True, italic=True, color="FF0000", size=14),
        fill=PatternFill(start_color="FFFF00", fill_type="solid"),
        border=Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="medium", color="0000FF"),
            bottom=Side(style="medium", color="0000FF")
        ),
        alignment=Alignment(horizontal="center", vertical="center"),
    )
    rule = Rule(type="cellIs", operator="greaterThan", formula=["50"], dxf=dxf)
    ws.conditional_formatting.add("A1:A10", rule)
    
    wb.save(os.path.join(OUTPUT_DIR, "dxf_full_style.xlsx"))

def create_dxf_font_effects_test():
    """Test: DXF with various font effects."""
    wb = Workbook()
    ws = wb.active
    ws.title = "FontEffects"
    
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=f"Value {i}")
    
    dxf = DifferentialStyle(
        font=Font(
            bold=True,
            italic=True,
            underline="single",
            strike=True,
            color="0000FF"
        )
    )
    rule = Rule(type="cellIs", operator="equal", formula=['"Value 3"'], dxf=dxf)
    ws.conditional_formatting.add("A1:A5", rule)
    
    wb.save(os.path.join(OUTPUT_DIR, "dxf_font_effects.xlsx"))

if __name__ == "__main__":
    print("Generating DXF test fixtures...")
    create_dxf_numfmt_test()
    create_dxf_alignment_test()
    create_dxf_protection_test()
    create_dxf_full_style_test()
    create_dxf_font_effects_test()
    print(f"Fixtures saved to {OUTPUT_DIR}")
```

### 5.2 Rust Integration Tests

Add to `tests/xlsx_formatting_roundtrip.rs`:

```rust
/// Test DXF with number format roundtrip
#[test]
fn test_roundtrip_dxf_number_format() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    
    for i in 0..5 {
        sheet.set_cell_value_at(i, 0, (i as f64 + 1.0) * 0.1).unwrap();
    }
    
    let style = Style::new()
        .fill_color(Color::rgb(255, 255, 0))
        .number_format("#,##0.00%");
    
    let rule = ConditionalFormatRule::cell_is_greater_than("0.3")
        .with_range(CellRange::parse("A1:A5").unwrap())
        .with_format(style);
    sheet.add_conditional_format(rule);
    
    // Roundtrip
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    
    // Verify number format preserved
    let sheet2 = wb2.worksheet(0).unwrap();
    let rules = sheet2.conditional_formats();
    assert_eq!(rules.len(), 1);
    
    if let Some(ref format) = rules[0].format {
        match &format.number_format {
            NumberFormat::Custom(code) => assert!(code.contains("%")),
            _ => panic!("Expected custom number format"),
        }
    }
}

/// Test DXF with alignment roundtrip
#[test]
fn test_roundtrip_dxf_alignment() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    
    sheet.set_cell_value("A1", "Test").unwrap();
    
    let mut style = Style::new();
    style.alignment = Alignment::new()
        .with_horizontal(HorizontalAlignment::Center)
        .with_vertical(VerticalAlignment::Center)
        .with_wrap(true)
        .with_rotation(45);
    
    let rule = ConditionalFormatRule::cell_is_equal_to("\"Test\"")
        .with_range(CellRange::parse("A1").unwrap())
        .with_format(style);
    sheet.add_conditional_format(rule);
    
    // Roundtrip
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    
    // Verify alignment preserved
    let sheet2 = wb2.worksheet(0).unwrap();
    let rules = sheet2.conditional_formats();
    
    if let Some(ref format) = rules[0].format {
        assert_eq!(format.alignment.horizontal, HorizontalAlignment::Center);
        assert_eq!(format.alignment.vertical, VerticalAlignment::Center);
        assert!(format.alignment.wrap_text);
        assert_eq!(format.alignment.rotation, 45);
    }
}

/// Test DXF with protection roundtrip
#[test]
fn test_roundtrip_dxf_protection() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    
    sheet.set_cell_value("A1", 500).unwrap();
    
    let mut style = Style::new();
    style.protection = Protection { locked: false, hidden: true };
    
    let rule = ConditionalFormatRule::cell_is_greater_than("100")
        .with_range(CellRange::parse("A1").unwrap())
        .with_format(style);
    sheet.add_conditional_format(rule);
    
    // Roundtrip
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    
    // Verify protection preserved
    let sheet2 = wb2.worksheet(0).unwrap();
    let rules = sheet2.conditional_formats();
    
    if let Some(ref format) = rules[0].format {
        assert!(!format.protection.locked);
        assert!(format.protection.hidden);
    }
}
```

### 5.3 Fixture Validation Tests

```rust
/// Test reading external DXF fixture (created by openpyxl)
#[test]
fn test_read_fixture_dxf_full_style() {
    let path = "tests/fixtures/dxf/dxf_full_style.xlsx";
    if !std::path::Path::new(path).exists() {
        eprintln!("Skipping test: fixture not found at {}", path);
        return;
    }
    
    let wb = XlsxReader::read_file(path).unwrap();
    let sheet = wb.worksheet(0).unwrap();
    
    // Should have one CF rule
    assert_eq!(sheet.conditional_format_count(), 1);
    
    let rules = sheet.conditional_formats();
    let format = rules[0].format.as_ref().expect("Rule should have format");
    
    // Verify font
    assert!(format.font.bold);
    assert!(format.font.italic);
    
    // Verify fill
    assert!(!matches!(format.fill, FillStyle::None));
    
    // Verify alignment
    assert_eq!(format.alignment.horizontal, HorizontalAlignment::Center);
}
```

---

## 6. API Design Decisions

### Decision 1: Reuse `Style` struct for DXF

**Rationale:**
- OOXML spec uses same element structure for DXF as regular XF
- Current implementation already uses `Style` for `rule.format`
- Simplicity: One unified style type
- DXF-specific behavior lives in writer/reader, not data model

**Implementation:**
- No changes to `Style` struct
- Writer intelligently omits defaults
- Reader parses same structure

### Decision 2: Inline numFmt in DXF

**Background:** In regular XF, number formats are referenced by ID. In DXF, they're written inline with both ID and format code.

**Implementation:**
```xml
<!-- Regular XF (in numFmts section): -->
<numFmt numFmtId="164" formatCode="#,##0.00"/>

<!-- DXF (inline): -->
<dxf>
  <numFmt numFmtId="164" formatCode="#,##0.00"/>
</dxf>
```

### Decision 3: Font Properties in DXF

Per XlsxWriter behavior:
- **Include if non-default:** bold, italic, strike, underline, color
- **Include only if explicitly set:** size, name (not required like in regular font)
- **DXF-specific:** condense, extend (for East Asian text)

---

## 7. Edge Cases & Quirks

### 7.1 Font in DXF vs Regular XF

| Property | Regular Font | DXF Font |
|----------|--------------|----------|
| `<sz>` | Required | Optional (omit if default) |
| `<name>` | Required | Optional (omit if default) |
| `<family>` | Optional | Typically omitted |
| `<scheme>` | Optional | Typically omitted |
| `<condense>` | Rare | Used for DXF |
| `<extend>` | Rare | Used for DXF |

### 7.2 Fill Color Semantics

From XlsxWriter:
```python
# In DXF, colors have different semantics:
if is_dxf_format:
    bg_color = xf_format.dxf_bg_color  # Different from regular bg_color!
    fg_color = xf_format.dxf_fg_color
```

**Our approach:** Use same `FillStyle` but handle solid fill specially:
- Solid fill in DXF: only fgColor, no bgColor needed

### 7.3 Border in DXF

From XlsxWriter:
```python
# Diagonal borders NOT allowed in DXF
if not is_dxf_format:
    self._write_sub_border("diagonal", ...)

# DXF adds empty vertical/horizontal
if is_dxf_format:
    self._write_sub_border("vertical", None, None)
    self._write_sub_border("horizontal", None, None)
```

### 7.4 PatternFill "none" in DXF

```python
# In DXF, pattern "none" handled differently:
if is_dxf_format and pattern <= 1:
    self._xml_start_tag("patternFill")  # No patternType attribute!
else:
    self._xml_start_tag("patternFill", [("patternType", patterns[pattern])])
```

### 7.5 Built-in Number Format IDs

| ID Range | Description |
|----------|-------------|
| 0 | General |
| 1-49 | Built-in formats |
| 50-163 | Reserved (locale-specific) |
| 164+ | Custom formats |

When writing DXF numFmt, use ID 164+ for any custom format.

---

## 8. Effort Estimates

### Phase 1: High Priority (Core Functionality)
| Task | Estimate | Dependencies |
|------|----------|--------------|
| 1.1 Number Format in DXF | 2-3 hours | None |
| 1.2 Alignment in DXF | 2-3 hours | None |
| 1.3 Test Fixtures | 2 hours | openpyxl |
| **Phase 1 Total** | **6-8 hours** | |

### Phase 2: Medium Priority
| Task | Estimate | Dependencies |
|------|----------|--------------|
| 2.1 Protection in DXF | 1-2 hours | None |
| 2.2 Gradient Fill | 3-4 hours | None |
| 2.3 Font Vertical Align | 1-2 hours | None |
| **Phase 2 Total** | **5-8 hours** | |

### Phase 3: Low Priority
| Task | Estimate | Dependencies |
|------|----------|--------------|
| 3.1 DXF Border Edges | 1 hour | None |
| 3.2 Font Extended Props | 2-3 hours | Add to FontStyle |
| 3.3 Extension List | 3-4 hours | New data structure |
| **Phase 3 Total** | **6-8 hours** | |

### Total Estimate: 17-24 hours

---

## Appendix A: XLSX Samples

### A.1 DXF with Full Style (expected output)

```xml
<dxfs count="1">
  <dxf>
    <font>
      <b/>
      <i/>
      <color rgb="FFFF0000"/>
    </font>
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFFFF00"/>
      </patternFill>
    </fill>
    <alignment horizontal="center" vertical="center" wrapText="1"/>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="medium"><color rgb="FF0000FF"/></top>
      <bottom style="medium"><color rgb="FF0000FF"/></bottom>
      <vertical/>
      <horizontal/>
    </border>
    <protection locked="0" hidden="1"/>
  </dxf>
</dxfs>
```

---

## Appendix B: References

1. **ECMA-376 5th Edition, Part 1**: §18.8.14 (dxf), §18.8.15 (dxfs)
2. **Microsoft OpenXML SDK**: `DifferentialFormat` class documentation
3. **XlsxWriter** (Python): `styles.py` - `_write_dxfs()`, `_write_font()`, `_write_fill()`, `_write_border()`
4. **openpyxl** (Python): `styles/differential.py`

---

## Appendix C: Checklist

### Phase 1 Checklist
- [ ] Implement `write_dxf_numfmt()` in styles.rs
- [ ] Add numFmt parsing in DXF reader
- [ ] Implement `write_dxf_alignment()` in styles.rs  
- [ ] Add alignment parsing in DXF reader
- [ ] Create Python fixture generation script
- [ ] Generate test fixtures
- [ ] Add Rust roundtrip tests
- [ ] Add fixture validation tests
- [ ] Run all tests and verify

### Phase 2 Checklist
- [ ] Implement `write_dxf_protection()` in styles.rs
- [ ] Add protection parsing in DXF reader
- [ ] Implement gradient fill writing
- [ ] Add gradient fill parsing in DXF reader
- [ ] Add font vertical align to DXF writer/reader
- [ ] Add tests for all Phase 2 features

### Phase 3 Checklist
- [ ] Add vertical/horizontal border elements to DXF writer
- [ ] Skip diagonal border in DXF
- [ ] Add font extended properties to FontStyle
- [ ] Implement font extended props in DXF writer/reader
- [ ] Design extension list preservation strategy
- [ ] Implement extLst reader/writer
- [ ] Add tests for all Phase 3 features
