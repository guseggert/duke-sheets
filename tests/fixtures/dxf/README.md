# DXF Test Fixtures

This directory contains test fixtures for DXF (Differential Format) support in duke-sheets.

## Fixture Sources

| Prefix | Source | Authority | Notes |
|--------|--------|-----------|-------|
| `pyuno_*.xlsx` | LibreOffice PyUNO | **Highest** | Created via LibreOffice's native UNO API |
| `dxf_*.xlsx` | openpyxl | Medium | Created via Python openpyxl library |
| `manual_*.xlsx` | Hand-crafted | Spec-compliant | Direct XML construction |

## Using the PyUNO Framework

### Build the Docker Image

```bash
docker build -t dxf-fixture-gen tests/fixtures/dxf/
```

### List Available Fixtures

```bash
docker run --rm dxf-fixture-gen --list
```

### Generate All Fixtures

```bash
docker run --rm -v $(pwd)/tests/fixtures/dxf:/output dxf-fixture-gen
```

### Generate Specific Fixtures

```bash
docker run --rm -v $(pwd)/tests/fixtures/dxf:/output dxf-fixture-gen pyuno_dxf_font.xlsx pyuno_dxf_fill.xlsx
```

## Creating Custom Fixtures

The framework uses a decorator-based approach. Add fixtures to `pyuno_framework.py`:

```python
from pyuno_framework import fixture, FixtureBuilder, StyleSpec

@fixture("my_custom_fixture.xlsx")
def my_fixture(b: FixtureBuilder):
    """Description of what this fixture tests."""
    b.set_sheet_name("TestSheet")
    
    # Add data
    b.add_data("A1", [10, 20, 30, 40, 50])
    
    # Add conditional format with style dict
    b.add_conditional_format(
        range="A1:A5",
        condition=("greater_than", 25),
        style={
            "bold": True,
            "fill_color": 0xFFFF00,  # Yellow
            "horizontal": "center",
        }
    )
    
    # Or use StyleSpec for full control
    b.add_conditional_format(
        range="B1:B5",
        condition=("equal", '"test"'),
        style=StyleSpec(
            bold=True,
            italic=True,
            font_color=0xFF0000,
            fill_color=0x00FF00,
            horizontal="center",
            vertical="center",
            wrap_text=True,
            border_style="thin",
            border_color=0x0000FF,
            number_format="0.00%",
        )
    )
```

## StyleSpec Properties

| Property | Type | Description |
|----------|------|-------------|
| `bold` | bool | Bold font |
| `italic` | bool | Italic font |
| `underline` | bool | Underlined text |
| `strikethrough` | bool | Strikethrough text |
| `font_color` | int | RGB color (e.g., `0xFF0000`) |
| `font_size` | float | Font size in points |
| `font_name` | str | Font family name |
| `fill_color` | int | Background RGB color |
| `horizontal` | str | `"left"`, `"center"`, `"right"` |
| `vertical` | str | `"top"`, `"center"`, `"bottom"` |
| `wrap_text` | bool | Enable text wrapping |
| `rotation` | int | Text rotation in degrees |
| `indent` | int | Indentation level |
| `border_style` | str | `"thin"`, `"medium"`, `"thick"` |
| `border_color` | int | Border RGB color |
| `number_format` | str | Format string (e.g., `"0.00%"`) |

## Condition Operators

- `"greater_than"` - Cell value > formula
- `"less_than"` - Cell value < formula  
- `"equal"` - Cell value = formula
- `"between"` - Cell value between two values

## Generated DXF Examples

### Font DXF
```xml
<dxf>
  <font>
    <b val="1"/>
    <color rgb="FFFF0000"/>
  </font>
</dxf>
```

### Alignment DXF
```xml
<dxf>
  <alignment horizontal="center" vertical="center" wrapText="true"/>
</dxf>
```

### Number Format DXF
```xml
<dxf>
  <numFmt numFmtId="164" formatCode="0.00%"/>
</dxf>
```

### Full Style DXF
```xml
<dxf>
  <font>
    <b val="1"/>
    <i val="1"/>
    <color rgb="FFFF0000"/>
  </font>
  <fill>
    <patternFill>
      <bgColor rgb="FFFFFF00"/>
    </patternFill>
  </fill>
  <alignment horizontal="center" vertical="center" wrapText="true"/>
  <border>
    <left style="dotted"/>
    <right style="dotted"/>
    <top style="dotted"/>
    <bottom style="dotted"/>
  </border>
</dxf>
```

## Regenerating openpyxl Fixtures

```bash
python3 tests/fixtures/dxf/generate_fixtures.py
```
