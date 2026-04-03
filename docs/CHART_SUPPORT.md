# Chart Support — OpenXML Type Checklist

> Auto-generated from [Open-XML-SDK `data/schemas/`](https://github.com/dotnet/Open-XML-SDK/tree/main/data/schemas).
> Each item references `schema_file#Types[index]` for traceability.
>
> Regenerate: `python3 tools/gen-chart-support.py`

**Total items: 840 checkboxes (292 implemented, 548 remaining), 48 abstract bases**

## 1. Package Structure

Relationships, content types, and part paths needed to embed
charts in an XLSX file. These are not in the XSD schemas —
they come from the SDK's part class definitions.

- [x] `xl/drawings/drawingN.xml` — DrawingsPart
- [x] `xl/charts/chartN.xml` — ChartPart
- [ ] `xl/charts/styleN.xml` — ChartStylePart (Office 2013+)
- [ ] `xl/charts/colorsN.xml` — ChartColorStylePart (Office 2013+)
- [ ] `xl/chartsheets/sheetN.xml` — ChartsheetPart
- [x] Worksheet → Drawing relationship (`RT_DRAWING`)
- [x] Drawing → Chart relationship (`RT_CHART`)
- [ ] Drawing → ChartEx relationship (`RT_CHART_EX`, Office 2016+)
- [ ] Chart → ChartStyle relationship
- [ ] Chart → ChartColorStyle relationship
- [x] `[Content_Types].xml` override for DrawingsPart
- [x] `[Content_Types].xml` override for ChartPart
- [ ] `[Content_Types].xml` override for ChartStylePart
- [ ] `[Content_Types].xml` override for ChartColorStylePart
- [x] `<drawing r:id="..."/>` element in worksheet XML

## 2. SpreadsheetDrawing (`xdr:` namespace)

Source: `schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json`

- [x] `xdr:twoCellAnchor` — **TwoCellAnchor**: Two Cell Anchor Shape Size (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[0]`)
- [x] `xdr:oneCellAnchor` — **OneCellAnchor**: One Cell Anchor Shape Size (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[1]`)
- [x] `xdr:absoluteAnchor` — **AbsoluteAnchor**: Absolute Anchor Shape Size (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[2]`)
- [ ] `xdr:sp` — **Shape**: Shape (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[3]`)
- [ ] `xdr:grpSp` — **GroupShape**: Group Shape (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[4]`)
- [x] `xdr:graphicFrame` — **GraphicFrame**: Graphic Frame (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[5]`)
- [ ] `xdr:cxnSp` — **ConnectionShape**: Connection Shape (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[6]`)
- [ ] `xdr:pic` — **Picture**: Defines the Picture Class (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[7]`)
- [ ] `xdr:contentPart` — **ContentPart**: Defines the ContentPart Class (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[8]`)
- [x] `xdr:wsDr` — **WorksheetDrawing**: Worksheet Drawing (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[9]`)
- [ ] `xdr:nvSpPr` — **NonVisualShapeProperties**: Non-Visual Properties for a Shape (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[10]`)
- [ ] `xdr:spPr` — **ShapeProperties**: Shape Properties (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[11]`)
- [ ] `xdr:style` — **ShapeStyle**: Defines the ShapeStyle Class (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[12]`)
- [ ] `xdr:txBody` — **TextBody**: Shape Text Body (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[13]`)
- [ ] `xdr:nvCxnSpPr` — **NonVisualConnectionShapeProperties**: Non-Visual Properties for a Connection Shape (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[14]`)
- [ ] `xdr:nvPicPr` — **NonVisualPictureProperties**: Non-Visual Properties for a Picture (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[15]`)
- [ ] `xdr:blipFill` — **BlipFill**: Picture Fill (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[16]`)
- [x] `xdr:nvGraphicFramePr` — **NonVisualGraphicFrameProperties**: Non-Visual Properties for a Graphic Frame (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[17]`)
- [x] `xdr:xfrm` — **Transform**: 2D Transform for Graphic Frames (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[18]`)
- [x] `xdr:col` — **ColumnId**: Column) (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[19]`)
- [x] `xdr:colOff` — **ColumnOffset**: Column Offset (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[20]`)
- [x] `xdr:rowOff` — **RowOffset**: Row Offset (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[21]`)
- [x] `xdr:row` — **RowId**: Row (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[22]`)
- [x] `xdr:from` — **FromMarker**: Starting Anchor Point (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[23]`)
- [x] `xdr:to` — **ToMarker**: Ending Anchor Point (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[24]`)
  - _abstract base: `MarkerType`_ — Defines the MarkerType Class (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[25]`)
- [x] `xdr:clientData` — **ClientData**: Client Data (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[26]`)
- [x] `xdr:ext` — **Extent**: Defines the Extent Class (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[27]`)
- [x] `xdr:pos` — **Position**: Position (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[28]`)
- [x] `xdr:cNvPr` — **NonVisualDrawingProperties**: Non-Visual Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[29]`)
- [ ] `xdr:cNvSpPr` — **NonVisualShapeDrawingProperties**: Connection Non-Visual Shape Properties (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[30]`)
- [ ] `xdr:cNvCxnSpPr` — **NonVisualConnectorShapeDrawingProperties**: Non-Visual Connector Shape Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[31]`)
- [ ] `xdr:cNvPicPr` — **NonVisualPictureDrawingProperties**: Non-Visual Picture Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[32]`)
- [x] `xdr:cNvGraphicFramePr` — **NonVisualGraphicFrameDrawingProperties**: Non-Visual Graphic Frame Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[33]`)
- [ ] `xdr:cNvGrpSpPr` — **NonVisualGroupShapeDrawingProperties**: Non-Visual Group Shape Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[34]`)
- [ ] `xdr:nvGrpSpPr` — **NonVisualGroupShapeProperties**: Non-Visual Properties for a Group Shape (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[35]`)
- [ ] `xdr:grpSpPr` — **GroupShapeProperties**: Group Shape Properties (`schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json#Types[36]`)

## 3. Chart Root & Structure (`c:` namespace)

Source: `schemas_openxmlformats_org_drawingml_2006_chart.json`

- [x] `c:chartSpace` — **ChartSpace**: Chart Space (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[101]`)
- [ ] `c:floor` — **Floor**: 3D floor formatting (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[306]`)
- [ ] `c:sideWall` — **SideWall**: 3D side wall formatting (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[307]`)
- [ ] `c:backWall` — **BackWall**: 3D back wall formatting (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[308]`)
  - _abstract base: `SurfaceType`_ — Defines the SurfaceType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[309]`)
- [x] `c:plotArea` — **PlotArea**: Plot data and formatting (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[310]`)
- [x] `c:chart` — **Chart**: Defines the Chart Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[319]`)

## 4. Chart Types (`c:` namespace)

- [x] `c:areaChart` — **AreaChart**: Area Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[183]`)
- [x] `c:area3DChart` — **Area3DChart**: 3D Area Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[184]`)
- [x] `c:lineChart` — **LineChart**: Line Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[185]`)
- [x] `c:line3DChart` — **Line3DChart**: 3D Line Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[186]`)
- [x] `c:stockChart` — **StockChart**: Stock Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[187]`)
- [x] `c:radarChart` — **RadarChart**: Radar Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[188]`)
- [x] `c:scatterChart` — **ScatterChart**: Scatter Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[189]`)
- [x] `c:pieChart` — **PieChart**: Pie Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[190]`)
- [x] `c:pie3DChart` — **Pie3DChart**: 3D Pie Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[191]`)
- [x] `c:doughnutChart` — **DoughnutChart**: Doughnut Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[192]`)
- [x] `c:barChart` — **BarChart**: Bar Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[193]`)
- [x] `c:bar3DChart` — **Bar3DChart**: 3D Bar Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[194]`)
- [ ] `c:ofPieChart` — **OfPieChart**: Pie of Pie or Bar of Pie Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[195]`)
- [x] `c:surfaceChart` — **SurfaceChart**: Surface Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[196]`)
- [x] `c:surface3DChart` — **Surface3DChart**: 3D Surface Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[197]`)
- [x] `c:bubbleChart` — **BubbleChart**: Bubble Charts (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[198]`)
- [x] `c:ext` — **StockChartExtension**: Defines the StockChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[208]`)
- [x] `c:ext` — **PieChartExtension**: Defines the PieChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[209]`)
- [x] `c:ext` — **Pie3DChartExtension**: Defines the Pie3DChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[210]`)
- [x] `c:ext` — **LineChartExtension**: Defines the LineChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[216]`)
- [x] `c:ext` — **Line3DChartExtension**: Defines the Line3DChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[217]`)
- [x] `c:ext` — **ScatterChartExtension**: Defines the ScatterChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[218]`)
- [x] `c:ext` — **RadarChartExtension**: Defines the RadarChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[219]`)
- [x] `c:ext` — **BarChartExtension**: Defines the BarChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[220]`)
- [x] `c:ext` — **Bar3DChartExtension**: Defines the Bar3DChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[221]`)
- [x] `c:ext` — **AreaChartExtension**: Defines the AreaChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[222]`)
- [x] `c:ext` — **Area3DChartExtension**: Defines the Area3DChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[223]`)
- [x] `c:ext` — **BubbleChartExtension**: Defines the BubbleChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[224]`)
- [x] `c:ext` — **SurfaceChartExtension**: Defines the SurfaceChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[225]`)
- [x] `c:ext` — **Surface3DChartExtension**: Defines the Surface3DChartExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[226]`)
- [x] `c:extLst` — **StockChartExtensionList**: Defines the StockChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[232]`)
- [x] `c:extLst` — **PieChartExtensionList**: Defines the PieChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[233]`)
- [x] `c:extLst` — **Pie3DChartExtensionList**: Defines the Pie3DChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[234]`)
- [x] `c:extLst` — **LineChartExtensionList**: Defines the LineChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[241]`)
- [x] `c:extLst` — **Line3DChartExtensionList**: Defines the Line3DChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[242]`)
- [x] `c:extLst` — **ScatterChartExtensionList**: Defines the ScatterChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[245]`)
- [x] `c:extLst` — **RadarChartExtensionList**: Defines the RadarChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[248]`)
- [x] `c:extLst` — **BarChartExtensionList**: Defines the BarChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[250]`)
- [x] `c:extLst` — **Bar3DChartExtensionList**: Defines the Bar3DChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[252]`)
- [x] `c:extLst` — **AreaChartExtensionList**: Defines the AreaChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[253]`)
- [x] `c:extLst` — **Area3DChartExtensionList**: Defines the Area3DChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[254]`)
- [x] `c:extLst` — **BubbleChartExtensionList**: Defines the BubbleChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[258]`)
- [x] `c:extLst` — **SurfaceChartExtensionList**: Defines the SurfaceChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[259]`)
- [x] `c:extLst` — **Surface3DChartExtensionList**: Defines the Surface3DChartExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[260]`)

## 5. Series Types (`c:` namespace)

- [x] `c:ser` — **LineChartSeries**: Defines the LineChartSeries Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[70]`)
- [x] `c:ser` — **BarChartSeries**: Bar Chart Series (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[74]`)
- [x] `c:ser` — **AreaChartSeries**: Area Chart Series (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[75]`)
- [x] `c:ser` — **PieChartSeries**: Pie Chart Series (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[76]`)
- [x] `c:ser` — **SurfaceChartSeries**: Surface Chart Series (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[77]`)
- [x] `c:ser` — **ScatterChartSeries**: Defines the ScatterChartSeries Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[244]`)
- [x] `c:ser` — **RadarChartSeries**: Defines the RadarChartSeries Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[247]`)
- [x] `c:ser` — **BubbleChartSeries**: Defines the BubbleChartSeries Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[255]`)
- [x] `c:extLst` — **LineSerExtensionList**: Defines the LineSerExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[287]`)
- [x] `c:ext` — **LineSerExtension**: Defines the LineSerExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[288]`)
- [x] `c:extLst` — **ScatterSerExtensionList**: Defines the ScatterSerExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[289]`)
- [x] `c:ext` — **ScatterSerExtension**: Defines the ScatterSerExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[290]`)
- [x] `c:extLst` — **RadarSerExtensionList**: Defines the RadarSerExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[291]`)
- [x] `c:ext` — **RadarSerExtension**: Defines the RadarSerExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[292]`)
- [x] `c:extLst` — **BarSerExtensionList**: Defines the BarSerExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[293]`)
- [x] `c:ext` — **BarSerExtension**: Defines the BarSerExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[294]`)
- [x] `c:extLst` — **AreaSerExtensionList**: Defines the AreaSerExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[295]`)
- [x] `c:ext` — **AreaSerExtension**: Defines the AreaSerExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[296]`)
- [x] `c:extLst` — **PieSerExtensionList**: Defines the PieSerExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[297]`)
- [x] `c:ext` — **PieSerExtension**: Defines the PieSerExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[298]`)
- [x] `c:extLst` — **BubbleSerExtensionList**: Defines the BubbleSerExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[299]`)
- [x] `c:ext` — **BubbleSerExtension**: Defines the BubbleSerExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[300]`)
- [x] `c:extLst` — **SurfaceSerExtensionList**: Defines the SurfaceSerExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[301]`)
- [x] `c:ext` — **SurfaceSerExtension**: Defines the SurfaceSerExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[302]`)

## 6. Data References (`c:` namespace)

- [x] `c:pt` — **NumericPoint**: Numeric Point (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[117]`)
- [x] `c:numRef` — **NumberReference**: Number Reference (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[119]`)
- [x] `c:numLit` — **NumberLiteral**: Number Literal (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[120]`)
- [x] `c:numCache` — **NumberingCache**: Defines the NumberingCache Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[121]`)
  - _abstract base: `NumberDataType`_ — Defines the NumberDataType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[122]`)
- [ ] `c:multiLvlStrRef` — **MultiLevelStringReference**: Multi Level String Reference (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[124]`)
- [x] `c:strRef` — **StringReference**: Defines the StringReference Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[125]`)
- [x] `c:strLit` — **StringLiteral**: String Literal (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[126]`)
- [x] `c:strCache` — **StringCache**: Defines the StringCache Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[127]`)
  - _abstract base: `StringDataType`_ — Defines the StringDataType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[128]`)
- [ ] `c:plus` — **Plus**: Plus (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[152]`)
- [ ] `c:minus` — **Minus**: Minus (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[153]`)
- [x] `c:val` — **Values**: Defines the Values Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[154]`)
- [x] `c:yVal` — **YValues**: Defines the YValues Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[155]`)
- [x] `c:bubbleSize` — **BubbleSize**: Defines the BubbleSize Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[156]`)
  - _abstract base: `NumberDataSourceType`_ — Defines the NumberDataSourceType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[157]`)
- [x] `c:pt` — **StringPoint**: String Point (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[206]`)
- [x] `c:ext` — **NumRefExtension**: Defines the NumRefExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[211]`)
- [x] `c:ext` — **StrDataExtension**: Defines the StrDataExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[212]`)
- [x] `c:ext` — **StrRefExtension**: Defines the StrRefExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[213]`)
- [x] `c:ext` — **MultiLvlStrRefExtension**: Defines the MultiLvlStrRefExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[214]`)
- [x] `c:extLst` — **NumRefExtensionList**: Defines the NumRefExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[235]`)
- [x] `c:extLst` — **StrDataExtensionList**: Defines the StrDataExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[236]`)
- [x] `c:extLst` — **StrRefExtensionList**: Defines the StrRefExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[237]`)
- [x] `c:extLst` — **MultiLvlStrRefExtensionList**: Defines the MultiLvlStrRefExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[239]`)
- [x] `c:cat` — **CategoryAxisData**: Defines the CategoryAxisData Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[284]`)
- [x] `c:xVal` — **XValues**: Defines the XValues Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[285]`)
  - _abstract base: `AxisDataSourceType`_ — Defines the AxisDataSourceType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[286]`)

## 7. Axes (`c:` namespace)

- [x] `c:scaling` — **Scaling**: Scaling (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[79]`)
- [x] `c:valAx` — **ValueAxis**: Value Axis (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[199]`)
- [x] `c:catAx` — **CategoryAxis**: Category Axis Data (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[200]`)
- [x] `c:dateAx` — **DateAxis**: Date Axis (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[201]`)
- [x] `c:serAx` — **SeriesAxis**: Series Axis (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[202]`)
- [x] `c:ext` — **CatAxExtension**: Defines the CatAxExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[227]`)
- [x] `c:ext` — **DateAxExtension**: Defines the DateAxExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[228]`)
- [x] `c:ext` — **SerAxExtension**: Defines the SerAxExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[229]`)
- [x] `c:ext` — **ValAxExtension**: Defines the ValAxExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[230]`)
- [x] `c:extLst` — **CatAxExtensionList**: Defines the CatAxExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[266]`)
- [x] `c:majorUnit` — **MajorUnit**: Defines the MajorUnit Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[271]`)
- [x] `c:minorUnit` — **MinorUnit**: Defines the MinorUnit Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[272]`)
  - _abstract base: `AxisUnitType`_ — Defines the AxisUnitType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[273]`)
- [x] `c:extLst` — **DateAxExtensionList**: Defines the DateAxExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[274]`)
- [x] `c:extLst` — **SerAxExtensionList**: Defines the SerAxExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[275]`)
- [x] `c:extLst` — **ValAxExtensionList**: Defines the ValAxExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[278]`)

## 8. Legend & Title (`c:` namespace)

- [x] `c:rich` — **RichText**: Rich Text (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[3]`)
- [x] `c:showLegendKey` — **ShowLegendKey**: Show Legend Key (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[6]`)
- [x] `c:tx` — **ChartText**: Defines the ChartText Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[51]`)
- [x] `c:tx` — **SeriesText**: Series Text (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[68]`)
- [x] `c:title` — **Title**: Title (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[81]`)
- [x] `c:legendPos` — **LegendPosition**: Legend Position (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[176]`)
- [ ] `c:legendEntry` — **LegendEntry**: Legend Entry (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[177]`)
- [x] `c:legend` — **Legend**: Legend data and formatting (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[311]`)

## 9. Data Labels & Data Points (`c:` namespace)

- [x] `c:dLblPos` — **DataLabelPosition**: Data Label Position (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[5]`)
- [x] `c:showDLblsOverMax` — **ShowDataLabelsOverMaximum**: True if we should render datalabels over the maximum scale (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[43]`)
- [x] `c:dLbls` — **DataLabels**: Data Labels (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[71]`)
- [ ] `c:dLbl` — **DataLabel**: Data Label (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[182]`)
- [x] `c:dPt` — **DataPoint**: Defines the DataPoint Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[281]`)

## 10. Trendlines & Error Bars (`c:` namespace)

- [x] `c:name` — **TrendlineName**: Trendline Name (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[48]`)
- [x] `c:val` — **ErrorBarValue**: Error Bar Value (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[95]`)
- [x] `c:trendlineType` — **TrendlineType**: Trendline Type (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[145]`)
- [ ] `c:trendlineLbl` — **TrendlineLabel**: Trendline Label (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[148]`)
- [x] `c:errBarType` — **ErrorBarType**: Error Bar Type (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[150]`)
- [x] `c:errValType` — **ErrorBarValueType**: Error Bar Value Type (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[151]`)
- [x] `c:trendline` — **Trendline**: Defines the Trendline Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[282]`)
- [x] `c:errBars` — **ErrorBars**: Defines the ErrorBars Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[283]`)

## 11. Formatting (`c:` namespace)

- [x] `c:spPr` — **ChartShapeProperties**: Defines the ChartShapeProperties Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[1]`)
- [ ] `c:txPr` — **TextProperties**: Defines the TextProperties Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[2]`)
- [x] `c:symbol` — **Symbol**: Symbol (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[141]`)
- [x] `c:size` — **Size**: Size (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[142]`)
- [x] `c:marker` — **Marker**: Marker (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[143]`)
- [ ] `c:pictureOptions` — **PictureOptions**: Defines the PictureOptions Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[144]`)
- [x] `c:spPr` — **ShapeProperties**: Defines the ShapeProperties Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[181]`)

## 12. Chart Configuration (`c:` namespace)

- [x] `c:layout` — **Layout**: Layout (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[50]`)
- [x] `c:leaderLines` — **LeaderLines**: Leader Lines (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[52]`)
- [x] `c:dropLines` — **DropLines**: Drop Lines (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[53]`)
- [x] `c:majorGridlines` — **MajorGridlines**: Major Gridlines (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[54]`)
- [x] `c:minorGridlines` — **MinorGridlines**: Minor Gridlines (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[55]`)
- [x] `c:serLines` — **SeriesLines**: Defines the SeriesLines Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[56]`)
- [x] `c:hiLowLines` — **HighLowLines**: Defines the HighLowLines Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[57]`)
  - _abstract base: `ChartLinesType`_ — Defines the ChartLinesType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[58]`)
- [x] `c:grouping` — **Grouping**: Grouping (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[69]`)
- [x] `c:barDir` — **BarDirection**: Bar Direction (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[72]`)
- [x] `c:grouping` — **BarGrouping**: Bar Grouping (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[73]`)
- [ ] `c:layoutTarget` — **LayoutTarget**: Layout Target (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[129]`)
- [ ] `c:xMode` — **LeftMode**: Left Mode (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[130]`)
- [ ] `c:yMode` — **TopMode**: Top Mode (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[131]`)
- [ ] `c:wMode` — **WidthMode**: Width Mode (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[132]`)
- [ ] `c:hMode` — **HeightMode**: Height Mode (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[133]`)
  - _abstract base: `LayoutModeType`_ — Defines the LayoutModeType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[134]`)
- [x] `c:manualLayout` — **ManualLayout**: Manual Layout (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[135]`)
- [x] `c:rotX` — **RotateX**: X Rotation (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[136]`)
- [x] `c:hPercent` — **HeightPercent**: Height Percent (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[137]`)
- [x] `c:rotY` — **RotateY**: Y Rotation (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[138]`)
- [x] `c:depthPercent` — **DepthPercent**: Depth Percent (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[139]`)
- [x] `c:perspective` — **Perspective**: Perspective (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[140]`)
- [x] `c:gapWidth` — **GapWidth**: Gap Width (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[158]`)
- [ ] `c:gapDepth` — **GapDepth**: Defines the GapDepth Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[159]`)
  - _abstract base: `GapAmountType`_ — Defines the GapAmountType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[160]`)
- [x] `c:upBars` — **UpBars**: Up Bars (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[161]`)
- [x] `c:downBars` — **DownBars**: Down Bars (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[162]`)
  - _abstract base: `UpDownBarType`_ — Defines the UpDownBarType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[163]`)
- [ ] `c:ofPieType` — **OfPieType**: Pie of Pie or Bar of Pie Type (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[164]`)
- [ ] `c:splitType` — **SplitType**: Split Type (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[165]`)
- [x] `c:orientation` — **Orientation**: Axis Orientation (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[174]`)
- [x] `c:upDownBars` — **UpDownBars**: Defines the UpDownBars Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[231]`)
- [x] `c:scatterStyle` — **ScatterStyle**: Defines the ScatterStyle Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[243]`)
- [x] `c:radarStyle` — **RadarStyle**: Defines the RadarStyle Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[246]`)
- [x] `c:overlap` — **Overlap**: Defines the Overlap Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[249]`)
- [ ] `c:shape` — **Shape**: Defines the Shape Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[251]`)
- [x] `c:bubbleScale` — **BubbleScale**: Defines the BubbleScale Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[256]`)
- [ ] `c:sizeRepresents` — **SizeRepresents**: Defines the SizeRepresents Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[257]`)
- [x] `c:view3D` — **View3D**: 3D view settings (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[305]`)
- [x] `c:dispBlanksAs` — **DisplayBlanksAs**: The way that blank cells are plotted on a chart (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[312]`)

## 13. Boolean Properties (`c:` namespace)

- [x] `c:showVal` — **ShowValue**: Show Value (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[7]`)
- [x] `c:showCatName` — **ShowCategoryName**: Show Category Name (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[8]`)
- [x] `c:showSerName` — **ShowSeriesName**: Show Series Name (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[9]`)
- [x] `c:showPercent` — **ShowPercent**: Show Percent (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[10]`)
- [x] `c:showBubbleSize` — **ShowBubbleSize**: Show Bubble Size (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[11]`)
- [x] `c:showLeaderLines` — **ShowLeaderLines**: Show Leader Lines (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[12]`)
- [x] `c:varyColors` — **VaryColors**: Defines the VaryColors Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[13]`)
- [x] `c:wireframe` — **Wireframe**: Wireframe (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[14]`)
- [x] `c:delete` — **Delete**: Delete (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[15]`)
- [x] `c:overlay` — **Overlay**: Overlay (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[16]`)
- [x] `c:rAngAx` — **RightAngleAxes**: Right Angle Axes (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[17]`)
- [x] `c:showHorzBorder` — **ShowHorizontalBorder**: Show Horizontal Border (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[18]`)
- [x] `c:showVertBorder` — **ShowVerticalBorder**: Show Vertical Border (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[19]`)
- [x] `c:showOutline` — **ShowOutlineBorder**: Show Outline Border (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[20]`)
- [x] `c:showKeys` — **ShowKeys**: Show Legend Keys (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[21]`)
- [x] `c:invertIfNegative` — **InvertIfNegative**: Invert if Negative (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[22]`)
- [ ] `c:bubble3D` — **Bubble3D**: 3D Bubble (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[23]`)
- [x] `c:dispRSqr` — **DisplayRSquaredValue**: Display R Squared Value (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[24]`)
- [x] `c:dispEq` — **DisplayEquation**: Display Equation (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[25]`)
- [x] `c:noEndCap` — **NoEndCap**: No End Cap (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[26]`)
- [ ] `c:applyToFront` — **ApplyToFront**: Apply To Front (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[27]`)
- [ ] `c:applyToSides` — **ApplyToSides**: Apply To Sides (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[28]`)
- [ ] `c:applyToEnd` — **ApplyToEnd**: Apply to End (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[29]`)
- [ ] `c:chartObject` — **ChartObject**: Chart Object (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[30]`)
- [ ] `c:data` — **Data**: Data Cannot Be Changed (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[31]`)
- [ ] `c:formatting` — **Formatting**: Formatting (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[32]`)
- [ ] `c:selection` — **Selection**: Selection (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[33]`)
- [ ] `c:userInterface` — **UserInterface**: User Interface (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[34]`)
- [ ] `c:autoUpdate` — **AutoUpdate**: Update Automatically (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[35]`)
- [x] `c:marker` — **ShowMarker**: Defines the ShowMarker Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[36]`)
- [x] `c:smooth` — **Smooth**: Defines the Smooth Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[37]`)
- [x] `c:showNegBubbles` — **ShowNegativeBubbles**: Defines the ShowNegativeBubbles Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[38]`)
- [ ] `c:auto` — **AutoLabeled**: Defines the AutoLabeled Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[39]`)
- [ ] `c:noMultiLvlLbl` — **NoMultiLevelLabels**: Defines the NoMultiLevelLabels Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[40]`)
- [x] `c:autoTitleDeleted` — **AutoTitleDeleted**: True if the chart automatic title has been deleted (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[41]`)
- [x] `c:plotVisOnly` — **PlotVisibleOnly**: True if only visible cells are plotted (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[42]`)
- [ ] `c:date1904` — **Date1904**: Defines the Date1904 Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[44]`)
- [x] `c:roundedCorners` — **RoundedCorners**: Defines the RoundedCorners Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[45]`)
  - _abstract base: `BooleanType`_ — Defines the BooleanType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[46]`)

## 14. Unsigned Integer Properties (`c:` namespace)

- [x] `c:idx` — **Index**: Index (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[59]`)
- [x] `c:order` — **Order**: Order (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[60]`)
- [x] `c:axId` — **AxisId**: Axis ID (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[61]`)
- [x] `c:crossAx` — **CrossingAxis**: Crossing Axis ID (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[62]`)
- [x] `c:ptCount` — **PointCount**: Point Count (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[63]`)
- [ ] `c:secondPiePt` — **SecondPiePoint**: Second Pie Point (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[64]`)
- [x] `c:explosion` — **Explosion**: Explosion (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[65]`)
- [ ] `c:fmtId` — **FormatId**: Format ID (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[66]`)
  - _abstract base: `UnsignedIntegerType`_ — Defines the UnsignedIntegerType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[67]`)

## 15. Print Settings & External Data (`c:` namespace)

- [ ] `c:headerFooter` — **HeaderFooter**: Header and Footer (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[178]`)
- [ ] `c:pageMargins` — **PageMargins**: Page Margins (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[179]`)
- [ ] `c:pageSetup` — **PageSetup**: Page Setup (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[180]`)
- [ ] `c:externalData` — **ExternalData**: Defines the ExternalData Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[320]`)
- [ ] `c:printSettings` — **PrintSettings**: Defines the PrintSettings Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[321]`)

## 16. Protection & Pivot (`c:` namespace)

- [ ] `c:pivotFmt` — **PivotFormat**: Pivot Format (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[175]`)
- [ ] `c:pivotFmts` — **PivotFormats**: pivot chart format persistence data (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[304]`)
- [ ] `c:pivotSource` — **PivotSource**: Defines the PivotSource Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[317]`)
- [ ] `c:protection` — **Protection**: Defines the Protection Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[318]`)

## 17. Extension Lists (`c:` namespace)

- [x] `c:extLst` — **ExtensionList**: Defines the ExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[118]`)
- [x] `c:ext` — **DLblsExtension**: Defines the DLblsExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[215]`)
- [x] `c:extLst` — **DLblsExtensionList**: Defines the DLblsExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[240]`)
- [x] `c:extLst` — **DLblExtensionList**: Defines the DLblExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[279]`)
- [x] `c:ext` — **DLblExtension**: Defines the DLblExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[280]`)
- [x] `c:extLst` — **ChartExtensionList**: Extensibility container (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[313]`)
- [x] `c:extLst` — **ChartSpaceExtensionList**: Defines the ChartSpaceExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[322]`)
- [x] `c:ext` — **ChartSpaceExtension**: Defines the ChartSpaceExtension Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[323]`)

## 18. Other (`c:` namespace)

- [x] `c:numFmt` — **NumberingFormat**: Number Format (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[0]`)
  - _abstract base: `TextBodyType`_ — Defines the TextBodyType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[4]`)
- [x] `c:separator` — **Separator**: Separator (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[47]`)
- [x] `c:f` — **Formula**: Defines the Formula Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[49]`)
- [ ] `c:bandFmts` — **BandFormats**: Band Formats (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[78]`)
- [x] `c:axPos` — **AxisPosition**: Axis Position (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[80]`)
- [x] `c:majorTickMark` — **MajorTickMark**: Major Tick Mark (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[82]`)
- [x] `c:minorTickMark` — **MinorTickMark**: Minor Tick Mark (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[83]`)
  - _abstract base: `TickMarkType`_ — Defines the TickMarkType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[84]`)
- [x] `c:tickLblPos` — **TickLabelPosition**: Tick Label Position (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[85]`)
- [x] `c:crosses` — **Crosses**: Crosses (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[86]`)
- [ ] `c:crossesAt` — **CrossesAt**: Crossing Value (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[87]`)
- [x] `c:x` — **Left**: Left (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[88]`)
- [x] `c:y` — **Top**: Top (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[89]`)
- [x] `c:w` — **Width**: Width (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[90]`)
- [x] `c:h` — **Height**: Height (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[91]`)
- [x] `c:forward` — **Forward**: Forward (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[92]`)
- [x] `c:backward` — **Backward**: Backward (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[93]`)
- [x] `c:intercept` — **Intercept**: Intercept (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[94]`)
- [ ] `c:splitPos` — **SplitPosition**: Split Position (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[96]`)
- [ ] `c:custUnit` — **CustomDisplayUnit**: Custom Display Unit (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[97]`)
- [x] `c:max` — **MaxAxisValue**: Maximum (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[98]`)
- [x] `c:min` — **MinAxisValue**: Minimum (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[99]`)
  - _abstract base: `DoubleType`_ — Defines the DoubleType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[100]`)
- [ ] `c:userShapes` — **UserShapes**: User Shapes (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[102]`)
- [x] `c:chart` — **ChartReference**: Reference to Chart Part (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[103]`)
- [ ] `c:legacyDrawingHF` — **LegacyDrawingHeaderFooter**: Legacy Drawing for Headers and Footers (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[104]`)
- [ ] `c:userShapes` — **UserShapesReference**: Defines the UserShapesReference Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[105]`)
  - _abstract base: `RelationshipIdType`_ — Defines the RelationshipIdType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[106]`)
- [x] `c:ext` — **Extension**: Extension (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[107]`)
- [x] `c:v` — **NumericValue**: Numeric Value (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[108]`)
- [ ] `c:formatCode` — **FormatCode**: Format Code (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[109]`)
- [ ] `c:oddHeader` — **OddHeader**: Odd Header (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[110]`)
- [ ] `c:oddFooter` — **OddFooter**: Odd Footer (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[111]`)
- [ ] `c:evenHeader` — **EvenHeader**: Even Header (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[112]`)
- [ ] `c:evenFooter` — **EvenFooter**: Even Footer (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[113]`)
- [ ] `c:firstHeader` — **FirstHeader**: First Header (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[114]`)
- [ ] `c:firstFooter` — **FirstFooter**: First Footer (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[115]`)
- [ ] `c:name` — **PivotTableName**: Pivot Name (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[116]`)
- [ ] `c:lvl` — **Level**: Level (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[123]`)
- [x] `c:order` — **PolynomialOrder**: Polynomial Trendline Order (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[146]`)
- [x] `c:period` — **Period**: Period (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[147]`)
- [x] `c:errDir` — **ErrorDirection**: Error Bar Direction (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[149]`)
- [ ] `c:custSplit` — **CustomSplit**: Custom Split (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[166]`)
- [ ] `c:secondPieSize` — **SecondPieSize**: Second Pie Size (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[167]`)
- [ ] `c:bandFmt` — **BandFormat**: Band Format (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[168]`)
- [ ] `c:pictureFormat` — **PictureFormat**: Picture Format (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[169]`)
- [ ] `c:pictureStackUnit` — **PictureStackUnit**: Picture Stack Unit (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[170]`)
- [ ] `c:builtInUnit` — **BuiltInUnit**: Built in Display Unit Value (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[171]`)
- [ ] `c:dispUnitsLbl` — **DisplayUnitsLabel**: Display Units Label (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[172]`)
- [ ] `c:logBase` — **LogBase**: Logarithmic Base (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[173]`)
- [x] `c:dTable` — **DataTable**: Data Table (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[203]`)
- [x] `c:firstSliceAng` — **FirstSliceAngle**: First Slice Angle (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[204]`)
- [x] `c:holeSize` — **HoleSize**: Hole Size (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[205]`)
- [ ] `c:thickness` — **Thickness**: Thickness (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[207]`)
- [ ] `c:multiLvlStrCache` — **MultiLevelStringCache**: Defines the MultiLevelStringCache Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[238]`)
- [ ] `c:lblAlgn` — **LabelAlignment**: Defines the LabelAlignment Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[261]`)
- [ ] `c:lblOffset` — **LabelOffset**: Defines the LabelOffset Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[262]`)
- [ ] `c:tickLblSkip` — **TickLabelSkip**: Defines the TickLabelSkip Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[263]`)
- [ ] `c:tickMarkSkip` — **TickMarkSkip**: Defines the TickMarkSkip Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[264]`)
  - _abstract base: `SkipType`_ — Defines the SkipType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[265]`)
- [ ] `c:baseTimeUnit` — **BaseTimeUnit**: Defines the BaseTimeUnit Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[267]`)
- [ ] `c:majorTimeUnit` — **MajorTimeUnit**: Defines the MajorTimeUnit Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[268]`)
- [ ] `c:minorTimeUnit` — **MinorTimeUnit**: Defines the MinorTimeUnit Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[269]`)
  - _abstract base: `TimeUnitType`_ — Defines the TimeUnitType Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[270]`)
- [x] `c:crossBetween` — **CrossBetween**: Defines the CrossBetween Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[276]`)
- [ ] `c:dispUnits` — **DisplayUnits**: Defines the DisplayUnits Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[277]`)
- [x] `c:ext` — **DataDisplayOptions16**: Defines the DataDisplayOptions16 Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[303]`)
- [ ] `c:lang` — **EditingLanguage**: Defines the EditingLanguage Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[314]`)
- [ ] `c:style` — **Style**: Defines the Style Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[315]`)
- [ ] `c:clrMapOvr` — **ColorMapOverride**: Defines the ColorMapOverride Class (`schemas_openxmlformats_org_drawingml_2006_chart.json#Types[316]`)

## 19. DrawingML Subset (`a:` namespace)

Only `a:` types directly referenced by chart or drawing elements.
Source: `schemas_openxmlformats_org_drawingml_2006_main.json`

- [ ] `a:tint` — **Tint**: Tint (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[7]`)
- [ ] `a:shade` — **Shade**: Shade (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[8]`)
- [ ] `a:alpha` — **Alpha**: Alpha (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[9]`)
  - _abstract base: `PositiveFixedPercentageType`_ — Defines the PositiveFixedPercentageType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[10]`)
- [ ] `a:comp` — **Complement**: Complement (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[11]`)
- [ ] `a:inv` — **Inverse**: Inverse (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[12]`)
- [ ] `a:gray` — **Gray**: Gray (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[13]`)
- [ ] `a:alphaOff` — **AlphaOffset**: Alpha Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[14]`)
- [ ] `a:alphaMod` — **AlphaModulation**: Alpha Modulation (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[15]`)
- [ ] `a:hueMod` — **HueModulation**: Hue Modulate (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[16]`)
  - _abstract base: `PositivePercentageType`_ — Defines the PositivePercentageType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[17]`)
- [ ] `a:hue` — **Hue**: Hue (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[18]`)
- [ ] `a:hueOff` — **HueOffset**: Hue Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[19]`)
- [ ] `a:sat` — **Saturation**: Saturation (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[20]`)
- [ ] `a:satOff` — **SaturationOffset**: Saturation Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[21]`)
- [ ] `a:satMod` — **SaturationModulation**: Saturation Modulation (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[22]`)
- [ ] `a:lum` — **Luminance**: Luminance (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[23]`)
- [ ] `a:lumOff` — **LuminanceOffset**: Luminance Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[24]`)
- [ ] `a:lumMod` — **LuminanceModulation**: Luminance Modulation (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[25]`)
- [ ] `a:red` — **Red**: Red (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[26]`)
- [ ] `a:redOff` — **RedOffset**: Red Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[27]`)
- [ ] `a:redMod` — **RedModulation**: Red Modulation (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[28]`)
- [ ] `a:green` — **Green**: Green (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[29]`)
- [ ] `a:greenOff` — **GreenOffset**: Green Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[30]`)
- [ ] `a:greenMod` — **GreenModulation**: Green Modification (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[31]`)
- [ ] `a:blue` — **Blue**: Blue (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[32]`)
- [ ] `a:blueOff` — **BlueOffset**: Blue Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[33]`)
- [ ] `a:blueMod` — **BlueModulation**: Blue Modification (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[34]`)
  - _abstract base: `PercentageType`_ — Defines the PercentageType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[35]`)
- [ ] `a:gamma` — **Gamma**: Gamma (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[36]`)
- [ ] `a:invGamma` — **InverseGamma**: Inverse Gamma (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[37]`)
- [x] `a:ext` — **Extension**: Extension (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[38]`)
- [ ] `a:scrgbClr` — **RgbColorModelPercentage**: RGB Color Model - Percentage Variant (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[39]`)
- [x] `a:srgbClr` — **RgbColorModelHex**: RGB Color Model - Hex Variant (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[40]`)
- [ ] `a:hslClr` — **HslColor**: Hue, Saturation, Luminance Color Model (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[41]`)
- [ ] `a:sysClr` — **SystemColor**: System Color (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[42]`)
- [ ] `a:schemeClr` — **SchemeColor**: Scheme Color (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[43]`)
- [ ] `a:prstClr` — **PresetColor**: Preset Color (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[44]`)
- [ ] `a:sp3d` — **Shape3DType**: Apply 3D shape properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[45]`)
- [ ] `a:flatTx` — **FlatText**: No text in 3D scene (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[46]`)
- [ ] `a:tile` — **Tile**: Tile (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[49]`)
- [ ] `a:stretch` — **Stretch**: Stretch (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[50]`)
- [x] `a:noFill` — **NoFill**: Defines the NoFill Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[51]`)
- [x] `a:solidFill` — **SolidFill**: Defines the SolidFill Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[52]`)
- [ ] `a:gradFill` — **GradientFill**: Defines the GradientFill Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[53]`)
- [ ] `a:blipFill` — **BlipFill**: Defines the BlipFill Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[54]`)
- [ ] `a:pattFill` — **PatternFill**: Pattern Fill (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[55]`)
- [ ] `a:grpFill` — **GroupFill**: Group Fill (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[56]`)
- [ ] `a:cont` — **EffectContainer**: Effect Container (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[57]`)
- [ ] `a:effectDag` — **EffectDag**: Effect Container (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[58]`)
  - _abstract base: `EffectContainerType`_ — Defines the EffectContainerType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[59]`)
- [ ] `a:effectLst` — **EffectList**: Effect Container (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[89]`)
- [ ] `a:custGeom` — **CustomGeometry**: Custom geometry (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[90]`)
- [ ] `a:prstGeom` — **PresetGeometry**: Preset geometry (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[91]`)
- [ ] `a:prstTxWarp` — **PresetTextWarp**: Preset Text Warp (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[92]`)
- [ ] `a:fillRef` — **FillReference**: Fill Reference (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[99]`)
- [ ] `a:effectRef` — **EffectReference**: Effect Reference (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[100]`)
- [ ] `a:lnRef` — **LineReference**: Defines the LineReference Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[101]`)
  - _abstract base: `StyleMatrixReferenceType`_ — Defines the StyleMatrixReferenceType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[102]`)
- [ ] `a:fontRef` — **FontReference**: Defines the FontReference Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[108]`)
- [ ] `a:noAutofit` — **NoAutoFit**: No AutoFit (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[109]`)
- [ ] `a:normAutofit` — **NormalAutoFit**: Normal AutoFit (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[110]`)
- [ ] `a:spAutoFit` — **ShapeAutoFit**: Shape AutoFit (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[111]`)
- [ ] `a:buClr` — **BulletColor**: Color Specified (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[113]`)
- [ ] `a:extrusionClr` — **ExtrusionColor**: Extrusion Color (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[114]`)
- [ ] `a:contourClr` — **ContourColor**: Contour Color (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[115]`)
- [ ] `a:clrFrom` — **ColorFrom**: Change Color From (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[116]`)
- [ ] `a:clrTo` — **ColorTo**: Change Color To (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[117]`)
- [ ] `a:fgClr` — **ForegroundColor**: Foreground color (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[118]`)
- [ ] `a:bgClr` — **BackgroundColor**: Background color (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[119]`)
- [ ] `a:highlight` — **Highlight**: Defines the Highlight Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[120]`)
  - _abstract base: `ColorType`_ — Defines the ColorType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[121]`)
- [ ] `a:buFont` — **BulletFont**: Specified (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[126]`)
- [ ] `a:latin` — **LatinFont**: Latin Font (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[127]`)
- [ ] `a:ea` — **EastAsianFont**: East Asian Font (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[128]`)
- [ ] `a:cs` — **ComplexScriptFont**: Complex Script Font (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[129]`)
- [ ] `a:sym` — **SymbolFont**: Defines the SymbolFont Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[130]`)
  - _abstract base: `TextFontType`_ — Defines the TextFontType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[131]`)
- [ ] `a:uLnTx` — **UnderlineFollowsText**: Underline Follows Text (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[136]`)
- [ ] `a:uLn` — **Underline**: Underline Stroke (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[137]`)
- [x] `a:ln` — **Outline**: Defines the Outline Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[138]`)
- [ ] `a:lnL` — **LeftBorderLineProperties**: Left Border Line Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[139]`)
- [ ] `a:lnR` — **RightBorderLineProperties**: Right Border Line Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[140]`)
- [ ] `a:lnT` — **TopBorderLineProperties**: Top Border Line Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[141]`)
- [ ] `a:lnB` — **BottomBorderLineProperties**: Bottom Border Line Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[142]`)
- [ ] `a:lnTlToBr` — **TopLeftToBottomRightBorderLineProperties**: Top-Left to Bottom-Right Border Line Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[143]`)
- [ ] `a:lnBlToTr` — **BottomLeftToTopRightBorderLineProperties**: Bottom-Left to Top-Right Border Line Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[144]`)
  - _abstract base: `LinePropertiesType`_ — Defines the LinePropertiesType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[145]`)
- [ ] `a:uFillTx` — **UnderlineFillText**: Underline Fill Properties Follow Text (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[146]`)
- [ ] `a:uFill` — **UnderlineFill**: Underline Fill (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[147]`)
- [ ] `a:graphic` — **Graphic**: Graphic Object (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[151]`)
- [ ] `a:blip` — **Blip**: Defines the Blip Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[152]`)
- [x] `a:extLst` — **ExtensionList**: Defines the ExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[160]`)
- [ ] `a:scene3d` — **Scene3DType**: 3D Scene Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[166]`)
- [ ] `a:off` — **Offset**: Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[191]`)
- [ ] `a:chOff` — **ChildOffset**: Child Offset (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[192]`)
  - _abstract base: `Point2DType`_ — Defines the Point2DType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[193]`)
- [ ] `a:ext` — **Extents**: Extents (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[194]`)
- [ ] `a:chExt` — **ChildExtents**: Child Extents (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[195]`)
  - _abstract base: `PositiveSize2DType`_ — Defines the PositiveSize2DType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[196]`)
- [ ] `a:spLocks` — **ShapeLocks**: Shape Locks (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[197]`)
- [ ] `a:cxnSpLocks` — **ConnectionShapeLocks**: Connection Shape Locks (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[198]`)
- [ ] `a:stCxn` — **StartConnection**: Connection Start (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[199]`)
- [ ] `a:endCxn` — **EndConnection**: Connection End (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[200]`)
  - _abstract base: `ConnectionType`_ — Defines the ConnectionType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[201]`)
- [ ] `a:graphicFrameLocks` — **GraphicFrameLocks**: Graphic Frame Locks (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[202]`)
- [ ] `a:txBody` — **TextBody**: Shape Text Body (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[208]`)
- [ ] `a:xfrm` — **Transform2D**: Defines the Transform2D Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[210]`)
- [ ] `a:cNvSpPr` — **NonVisualShapeDrawingProperties**: Non-Visual Shape Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[212]`)
- [ ] `a:spPr` — **ShapeProperties**: Visual Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[214]`)
- [ ] `a:style` — **ShapeStyle**: Style (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[216]`)
- [ ] `a:cNvCxnSpPr` — **NonVisualConnectorShapeDrawingProperties**: Non-Visual Connector Shape Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[217]`)
- [ ] `a:cNvPicPr` — **NonVisualPictureDrawingProperties**: Non-Visual Picture Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[219]`)
- [x] `a:cNvGraphicFramePr` — **NonVisualGraphicFrameDrawingProperties**: Non-Visual Graphic Frame Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[221]`)
- [ ] `a:cNvGrpSpPr` — **NonVisualGroupShapeDrawingProperties**: Non-Visual Group Shape Drawing Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[223]`)
- [ ] `a:fillToRect` — **FillToRectangle**: Fill To Rectangle (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[236]`)
- [ ] `a:tileRect` — **TileRectangle**: Tile Rectangle (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[237]`)
- [ ] `a:fillRect` — **FillRectangle**: Fill Rectangle (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[238]`)
- [ ] `a:srcRect` — **SourceRectangle**: Source Rectangle (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[239]`)
  - _abstract base: `RelativeRectangleType`_ — Defines the RelativeRectangleType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[240]`)
- [ ] `a:xfrm` — **TransformGroup**: 2D Transform for Grouped Objects (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[265]`)
- [ ] `a:bodyPr` — **BodyProperties**: Defines the BodyProperties Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[266]`)
- [ ] `a:lstStyle` — **ListStyle**: Defines the ListStyle Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[267]`)
- [ ] `a:overrideClrMapping` — **OverrideColorMapping**: Override Color Mapping (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[272]`)
- [ ] `a:clrMap` — **ColorMap**: Defines the ColorMap Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[273]`)
  - _abstract base: `ColorMappingType`_ — Defines the ColorMappingType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[274]`)
- [ ] `a:endParaRPr` — **EndParagraphRunProperties**: End Paragraph Run Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[327]`)
- [ ] `a:rPr` — **RunProperties**: Text Run Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[328]`)
- [ ] `a:defRPr` — **DefaultRunProperties**: Default Text Run Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[329]`)
  - _abstract base: `TextCharacterPropertiesType`_ — Defines the TextCharacterPropertiesType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[330]`)
- [ ] `a:p` — **Paragraph**: Text Paragraphs (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[331]`)
- [ ] `a:extLst` — **ShapePropertiesExtensionList**: Defines the ShapePropertiesExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[343]`)
- [ ] `a:grpSpPr` — **VisualGroupShapeProperties**: Visual Group Shape Properties (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[345]`)
- [ ] `a:grpSpLocks` — **GroupShapeLocks**: Defines the GroupShapeLocks Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[355]`)
- [ ] `a:extLst` — **NonVisualGroupDrawingShapePropsExtensionList**: Defines the NonVisualGroupDrawingShapePropsExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[356]`)
- [ ] `a:hlinkClick` — **HyperlinkOnClick**: Defines the HyperlinkOnClick Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[361]`)
- [ ] `a:hlinkMouseOver` — **HyperlinkOnMouseOver**: Defines the HyperlinkOnMouseOver Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[362]`)
- [ ] `a:hlinkHover` — **HyperlinkOnHover**: Defines the HyperlinkOnHover Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[363]`)
  - _abstract base: `HyperlinkType`_ — Defines the HyperlinkType Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[364]`)
- [ ] `a:rtl` — **RightToLeft**: Defines the RightToLeft Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[365]`)
- [ ] `a:extLst` — **NonVisualDrawingPropertiesExtensionList**: Defines the NonVisualDrawingPropertiesExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[366]`)
- [ ] `a:picLocks` — **PictureLocks**: Defines the PictureLocks Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[378]`)
- [ ] `a:extLst` — **NonVisualPicturePropertiesExtensionList**: Defines the NonVisualPicturePropertiesExtensionList Class (`schemas_openxmlformats_org_drawingml_2006_main.json#Types[379]`)

## 20. Extensions: `c14:` namespace

Namespace: `http://schemas.microsoft.com/office/drawing/2007/8/2/chart`
Source: `schemas_microsoft_com_office_drawing_2007_8_2_chart.json`

- [ ] `c14:pivotOptions` — **PivotOptions**: Defines the PivotOptions Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[0]`)
- [ ] `c14:sketchOptions` — **SketchOptions**: Defines the SketchOptions Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[1]`)
- [ ] `c14:invertSolidFillFmt` — **InvertSolidFillFormat**: Defines the InvertSolidFillFormat Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[2]`)
- [ ] `c14:style` — **Style**: Defines the Style Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[3]`)
- [ ] `c14:spPr` — **ShapeProperties**: Defines the ShapeProperties Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[4]`)
- [ ] `c14:dropZoneFilter` — **DropZoneFilter**: Defines the DropZoneFilter Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[5]`)
- [ ] `c14:dropZoneCategories` — **DropZoneCategories**: Defines the DropZoneCategories Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[6]`)
- [ ] `c14:dropZoneData` — **DropZoneData**: Defines the DropZoneData Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[7]`)
- [ ] `c14:dropZoneSeries` — **DropZoneSeries**: Defines the DropZoneSeries Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[8]`)
- [ ] `c14:dropZonesVisible` — **DropZonesVisible**: Defines the DropZonesVisible Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[9]`)
- [ ] `c14:inSketchMode` — **InSketchMode**: Defines the InSketchMode Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[10]`)
  - _abstract base: `BooleanFalseType`_ — Defines the BooleanFalseType Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[11]`)
- [ ] `c14:showSketchBtn` — **ShowSketchButton**: Defines the ShowSketchButton Class (`schemas_microsoft_com_office_drawing_2007_8_2_chart.json#Types[12]`)

## 21. Extensions: `c15:` namespace

Namespace: `http://schemas.microsoft.com/office/drawing/2012/chart`
Source: `schemas_microsoft_com_office_drawing_2012_chart.json`

- [ ] `c15:pivotSource` — **PivotSource**: Defines the PivotSource Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[0]`)
- [x] `c15:numFmt` — **NumberingFormat**: Defines the NumberingFormat Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[1]`)
- [ ] `c15:spPr` — **ShapeProperties**: Defines the ShapeProperties Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[2]`)
- [x] `c15:layout` — **Layout**: Defines the Layout Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[3]`)
- [ ] `c15:fullRef` — **FullReference**: Defines the FullReference Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[4]`)
- [ ] `c15:levelRef` — **LevelReference**: Defines the LevelReference Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[5]`)
- [ ] `c15:formulaRef` — **FormulaReference**: Defines the FormulaReference Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[6]`)
- [ ] `c15:filteredSeriesTitle` — **FilteredSeriesTitle**: Defines the FilteredSeriesTitle Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[7]`)
- [ ] `c15:filteredCategoryTitle` — **FilteredCategoryTitle**: Defines the FilteredCategoryTitle Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[8]`)
- [ ] `c15:filteredAreaSeries` — **FilteredAreaSeries**: Defines the FilteredAreaSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[9]`)
- [ ] `c15:filteredBarSeries` — **FilteredBarSeries**: Defines the FilteredBarSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[10]`)
- [ ] `c15:filteredBubbleSeries` — **FilteredBubbleSeries**: Defines the FilteredBubbleSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[11]`)
- [ ] `c15:filteredLineSeries` — **FilteredLineSeriesExtension**: Defines the FilteredLineSeriesExtension Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[12]`)
- [ ] `c15:filteredPieSeries` — **FilteredPieSeries**: Defines the FilteredPieSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[13]`)
- [ ] `c15:filteredRadarSeries` — **FilteredRadarSeries**: Defines the FilteredRadarSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[14]`)
- [ ] `c15:filteredScatterSeries` — **FilteredScatterSeries**: Defines the FilteredScatterSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[15]`)
- [ ] `c15:filteredSurfaceSeries` — **FilteredSurfaceSeries**: Defines the FilteredSurfaceSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[16]`)
- [ ] `c15:datalabelsRange` — **DataLabelsRange**: Defines the DataLabelsRange Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[17]`)
- [ ] `c15:categoryFilterExceptions` — **CategoryFilterExceptions**: Defines the CategoryFilterExceptions Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[18]`)
- [ ] `c15:dlblFieldTable` — **DataLabelFieldTable**: Defines the DataLabelFieldTable Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[19]`)
- [ ] `c15:xForSave` — **ExceptionForSave**: Defines the ExceptionForSave Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[20]`)
- [ ] `c15:showDataLabelsRange` — **ShowDataLabelsRange**: Defines the ShowDataLabelsRange Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[21]`)
- [x] `c15:showLeaderLines` — **ShowLeaderLines**: Defines the ShowLeaderLines Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[22]`)
- [ ] `c15:autoCat` — **AutoGeneneratedCategories**: Defines the AutoGeneneratedCategories Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[23]`)
- [ ] `c15:invertIfNegative` — **InvertIfNegativeBoolean**: Defines the InvertIfNegativeBoolean Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[24]`)
- [ ] `c15:bubble3D` — **Bubble3D**: Defines the Bubble3D Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[25]`)
  - _abstract base: `BooleanType`_ — Defines the BooleanType Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[26]`)
- [x] `c15:tx` — **ChartText**: Defines the ChartText Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[27]`)
- [x] `c15:leaderLines` — **LeaderLines**: Defines the LeaderLines Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[28]`)
- [ ] `c15:sqref` — **SequenceOfReferences**: Defines the SequenceOfReferences Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[29]`)
- [x] `c15:f` — **Formula**: Defines the Formula Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[30]`)
- [ ] `c15:txfldGUID` — **TextFieldGuid**: Defines the TextFieldGuid Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[31]`)
- [ ] `c15:cat` — **AxisDataSourceType**: Defines the AxisDataSourceType Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[32]`)
- [x] `c15:ser` — **BarChartSeries**: Defines the BarChartSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[33]`)
- [x] `c15:ser` — **LineChartSeries**: Defines the LineChartSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[34]`)
- [x] `c15:ser` — **ScatterChartSeries**: Defines the ScatterChartSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[35]`)
- [x] `c15:ser` — **AreaChartSeries**: Defines the AreaChartSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[36]`)
- [x] `c15:ser` — **PieChartSeries**: Defines the PieChartSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[37]`)
- [x] `c15:ser` — **BubbleChartSeries**: Defines the BubbleChartSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[38]`)
- [x] `c15:ser` — **RadarChartSeries**: Defines the RadarChartSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[39]`)
- [x] `c15:ser` — **SurfaceChartSeries**: Defines the SurfaceChartSeries Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[40]`)
- [ ] `c15:dlblRangeCache` — **DataLabelsRangeChache**: Defines the DataLabelsRangeChache Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[41]`)
- [ ] `c15:dlblFieldTableCache` — **DataLabelFieldTableCache**: Defines the DataLabelFieldTableCache Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[42]`)
  - _abstract base: `StringDataType`_ — Defines the StringDataType Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[43]`)
- [x] `c15:explosion` — **Explosion**: Defines the Explosion Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[44]`)
- [x] `c15:marker` — **Marker**: Defines the Marker Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[45]`)
- [ ] `c15:dLbl` — **DataLabel**: Defines the DataLabel Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[46]`)
- [ ] `c15:categoryFilterException` — **CategoryFilterException**: Defines the CategoryFilterException Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[47]`)
- [ ] `c15:dlblFTEntry` — **DataLabelFieldTableEntry**: Defines the DataLabelFieldTableEntry Class (`schemas_microsoft_com_office_drawing_2012_chart.json#Types[48]`)

## 22. Extensions: `c16:` namespace

Namespace: `http://schemas.microsoft.com/office/drawing/2014/chart`
Source: `schemas_microsoft_com_office_drawing_2014_chart.json`

- [ ] `c16:spPr` — **ShapeProperties**: Defines the ShapeProperties Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[0]`)
- [ ] `c16:explosion` — **UnsignedIntegerType**: Defines the UnsignedIntegerType Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[1]`)
- [ ] `c16:invertIfNegative` — **InvertIfNegativeBoolean**: Defines the InvertIfNegativeBoolean Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[2]`)
- [ ] `c16:bubble3D` — **Bubble3DBoolean**: Defines the Bubble3DBoolean Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[3]`)
  - _abstract base: `BooleanType`_ — Defines the BooleanType Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[4]`)
- [x] `c16:marker` — **Marker**: Defines the Marker Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[5]`)
- [ ] `c16:dLbl` — **DLbl**: Defines the DLbl Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[6]`)
- [ ] `c16:categoryFilterExceptions` — **CategoryFilterExceptions**: Defines the CategoryFilterExceptions Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[7]`)
- [ ] `c16:pivotOptions16` — **PivotOptions16**: Defines the PivotOptions16 Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[8]`)
- [ ] `c16:datapointuniqueidmap` — **ChartDataPointUniqueIDMap**: Defines the ChartDataPointUniqueIDMap Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[9]`)
- [ ] `c16:uniqueId` — **UniqueIdChartUniqueID**: Defines the UniqueIdChartUniqueID Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[10]`)
- [ ] `c16:uniqueID` — **UniqueID**: Defines the UniqueID Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[11]`)
  - _abstract base: `UniqueIDChart`_ — Defines the UniqueIDChart Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[12]`)
- [ ] `c16:categoryFilterException` — **CategoryFilterException**: Defines the CategoryFilterException Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[13]`)
- [ ] `c16:numCache` — **NumberDataType**: Defines the NumberDataType Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[14]`)
- [ ] `c16:filteredLitCache` — **NumFilteredLiteralCache**: Defines the NumFilteredLiteralCache Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[15]`)
- [ ] `c16:strCache` — **StringDataType**: Defines the StringDataType Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[16]`)
- [ ] `c16:filteredLitCache` — **StrFilteredLiteralCache**: Defines the StrFilteredLiteralCache Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[17]`)
- [ ] `c16:multiLvlStrCache` — **MultiLvlStrData**: Defines the MultiLvlStrData Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[18]`)
- [ ] `c16:filteredLitCache` — **MultiLvlStrFilteredLiteralCache**: Defines the MultiLvlStrFilteredLiteralCache Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[19]`)
- [ ] `c16:literalDataChart` — **LiteralDataChart**: Defines the LiteralDataChart Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[20]`)
- [ ] `c16:showExpandCollapseFieldButtons` — **BooleanFalse**: Defines the BooleanFalse Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[21]`)
- [ ] `c16:ptidx` — **XsdunsignedInt**: Defines the XsdunsignedInt Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[22]`)
- [ ] `c16:ptentry` — **ChartDataPointUniqueIDMapEntry**: Defines the ChartDataPointUniqueIDMapEntry Class (`schemas_microsoft_com_office_drawing_2014_chart.json#Types[23]`)

## 23. Extensions: `cx:` namespace

Namespace: `http://schemas.microsoft.com/office/drawing/2014/chartex`
Source: `schemas_microsoft_com_office_drawing_2014_chartex.json`

- [x] `cx:chartSpace` — **ChartSpace**: Defines the ChartSpace Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[0]`)
- [ ] `cx:chart` — **RelId**: Defines the RelId Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[1]`)
- [ ] `cx:openxmlsdk_49BECFFA_3B03_4D13_8272_D6CCB22579E3` — **Openxmlsdk_49BECFFA_3B03_4D13_8272_D6CCB22579E3XsdunsignedInt**: Defines the Openxmlsdk_49BECFFA_3B03_4D13_8272_D6CCB22579E3XsdunsignedInt Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[2]`)
- [ ] `cx:binCount` — **BinCountXsdunsignedInt**: Defines the BinCountXsdunsignedInt Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[3]`)
- [ ] `cx:ext` — **Extension2**: Defines the Extension2 Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[4]`)
- [ ] `cx:minColor` — **MinColorSolidColorFillProperties**: Defines the MinColorSolidColorFillProperties Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[5]`)
- [ ] `cx:midColor` — **MidColorSolidColorFillProperties**: Defines the MidColorSolidColorFillProperties Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[6]`)
- [ ] `cx:maxColor` — **MaxColorSolidColorFillProperties**: Defines the MaxColorSolidColorFillProperties Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[7]`)
  - _abstract base: `OpenXmlSolidColorFillPropertiesElement`_ — Defines the OpenXmlSolidColorFillPropertiesElement Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[8]`)
- [ ] `cx:pt` — **ChartStringValue**: Defines the ChartStringValue Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[9]`)
- [x] `cx:f` — **Formula**: Defines the Formula Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[10]`)
- [ ] `cx:nf` — **NfFormula**: Defines the NfFormula Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[11]`)
  - _abstract base: `OpenXmlFormulaElement`_ — Defines the OpenXmlFormulaElement Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[12]`)
- [ ] `cx:lvl` — **StringLevel**: Defines the StringLevel Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[13]`)
- [x] `cx:pt` — **NumericValue**: Defines the NumericValue Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[14]`)
- [ ] `cx:lvl` — **NumericLevel**: Defines the NumericLevel Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[15]`)
- [ ] `cx:numDim` — **NumericDimension**: Defines the NumericDimension Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[16]`)
- [ ] `cx:strDim` — **StringDimension**: Defines the StringDimension Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[17]`)
- [x] `cx:extLst` — **ExtensionList**: Defines the ExtensionList Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[18]`)
- [ ] `cx:externalData` — **ExternalData**: Defines the ExternalData Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[19]`)
- [ ] `cx:data` — **Data**: Defines the Data Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[20]`)
- [ ] `cx:v` — **VXsdstring**: Defines the VXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[21]`)
- [ ] `cx:copyright` — **CopyrightXsdstring**: Defines the CopyrightXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[22]`)
- [ ] `cx:separator` — **SeparatorXsdstring**: Defines the SeparatorXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[23]`)
- [ ] `cx:oddHeader` — **OddHeaderXsdstring**: Defines the OddHeaderXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[24]`)
- [ ] `cx:oddFooter` — **OddFooterXsdstring**: Defines the OddFooterXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[25]`)
- [ ] `cx:evenHeader` — **EvenHeaderXsdstring**: Defines the EvenHeaderXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[26]`)
- [ ] `cx:evenFooter` — **EvenFooterXsdstring**: Defines the EvenFooterXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[27]`)
- [ ] `cx:firstHeader` — **FirstHeaderXsdstring**: Defines the FirstHeaderXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[28]`)
- [ ] `cx:firstFooter` — **FirstFooterXsdstring**: Defines the FirstFooterXsdstring Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[29]`)
- [ ] `cx:txData` — **TextData**: Defines the TextData Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[30]`)
- [ ] `cx:rich` — **RichTextBody**: Defines the RichTextBody Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[31]`)
- [ ] `cx:txPr` — **TxPrTextBody**: Defines the TxPrTextBody Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[32]`)
  - _abstract base: `TextBodyType`_ — Defines the TextBodyType Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[33]`)
- [ ] `cx:tx` — **Text**: Defines the Text Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[34]`)
- [ ] `cx:spPr` — **ShapeProperties**: Defines the ShapeProperties Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[35]`)
- [ ] `cx:offset` — **Offset**: Defines the Offset Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[36]`)
- [ ] `cx:unitsLabel` — **AxisUnitsLabel**: Defines the AxisUnitsLabel Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[37]`)
- [ ] `cx:catScaling` — **CategoryAxisScaling**: Defines the CategoryAxisScaling Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[38]`)
- [ ] `cx:valScaling` — **ValueAxisScaling**: Defines the ValueAxisScaling Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[39]`)
- [ ] `cx:title` — **AxisTitle**: Defines the AxisTitle Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[40]`)
- [ ] `cx:units` — **AxisUnits**: Defines the AxisUnits Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[41]`)
- [ ] `cx:majorGridlines` — **MajorGridlinesGridlines**: Defines the MajorGridlinesGridlines Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[42]`)
- [ ] `cx:minorGridlines` — **MinorGridlinesGridlines**: Defines the MinorGridlinesGridlines Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[43]`)
  - _abstract base: `OpenXmlGridlinesElement`_ — Defines the OpenXmlGridlinesElement Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[44]`)
- [ ] `cx:majorTickMarks` — **MajorTickMarksTickMarks**: Defines the MajorTickMarksTickMarks Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[45]`)
- [ ] `cx:minorTickMarks` — **MinorTickMarksTickMarks**: Defines the MinorTickMarksTickMarks Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[46]`)
  - _abstract base: `OpenXmlTickMarksElement`_ — Defines the OpenXmlTickMarksElement Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[47]`)
- [ ] `cx:tickLabels` — **TickLabels**: Defines the TickLabels Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[48]`)
- [ ] `cx:numFmt` — **NumberFormat**: Defines the NumberFormat Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[49]`)
- [ ] `cx:binSize` — **Xsddouble**: Defines the Xsddouble Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[50]`)
- [ ] `cx:address` — **Address**: Defines the Address Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[51]`)
- [ ] `cx:geoLocation` — **GeoLocation**: Defines the GeoLocation Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[52]`)
- [ ] `cx:geoLocationQuery` — **GeoLocationQuery**: Defines the GeoLocationQuery Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[53]`)
- [ ] `cx:geoLocations` — **GeoLocations**: Defines the GeoLocations Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[54]`)
- [ ] `cx:geoLocationQueryResult` — **GeoLocationQueryResult**: Defines the GeoLocationQueryResult Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[55]`)
- [ ] `cx:geoPolygon` — **GeoPolygon**: Defines the GeoPolygon Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[56]`)
- [ ] `cx:geoPolygons` — **GeoPolygons**: Defines the GeoPolygons Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[57]`)
- [ ] `cx:copyrights` — **Copyrights**: Defines the Copyrights Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[58]`)
- [ ] `cx:geoDataEntityQuery` — **GeoDataEntityQuery**: Defines the GeoDataEntityQuery Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[59]`)
- [ ] `cx:geoData` — **GeoData**: Defines the GeoData Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[60]`)
- [ ] `cx:geoDataEntityQueryResult` — **GeoDataEntityQueryResult**: Defines the GeoDataEntityQueryResult Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[61]`)
- [ ] `cx:geoDataPointQuery` — **GeoDataPointQuery**: Defines the GeoDataPointQuery Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[62]`)
- [ ] `cx:geoDataPointToEntityQuery` — **GeoDataPointToEntityQuery**: Defines the GeoDataPointToEntityQuery Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[63]`)
- [ ] `cx:geoDataPointToEntityQueryResult` — **GeoDataPointToEntityQueryResult**: Defines the GeoDataPointToEntityQueryResult Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[64]`)
- [ ] `cx:entityType` — **EntityType**: Defines the EntityType Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[65]`)
- [ ] `cx:geoChildTypes` — **GeoChildTypes**: Defines the GeoChildTypes Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[66]`)
- [ ] `cx:geoHierarchyEntity` — **GeoHierarchyEntity**: Defines the GeoHierarchyEntity Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[67]`)
- [ ] `cx:geoChildEntitiesQuery` — **GeoChildEntitiesQuery**: Defines the GeoChildEntitiesQuery Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[68]`)
- [ ] `cx:geoChildEntities` — **GeoChildEntities**: Defines the GeoChildEntities Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[69]`)
- [ ] `cx:geoChildEntitiesQueryResult` — **GeoChildEntitiesQueryResult**: Defines the GeoChildEntitiesQueryResult Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[70]`)
- [ ] `cx:geoParentEntitiesQuery` — **GeoParentEntitiesQuery**: Defines the GeoParentEntitiesQuery Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[71]`)
- [ ] `cx:geoEntity` — **GeoEntity**: Defines the GeoEntity Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[72]`)
- [ ] `cx:geoParentEntity` — **GeoParentEntity**: Defines the GeoParentEntity Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[73]`)
- [ ] `cx:geoParentEntitiesQueryResult` — **GeoParentEntitiesQueryResult**: Defines the GeoParentEntitiesQueryResult Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[74]`)
- [ ] `cx:geoLocationQueryResults` — **GeoLocationQueryResults**: Defines the GeoLocationQueryResults Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[75]`)
- [ ] `cx:geoDataEntityQueryResults` — **GeoDataEntityQueryResults**: Defines the GeoDataEntityQueryResults Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[76]`)
- [ ] `cx:geoDataPointToEntityQueryResults` — **GeoDataPointToEntityQueryResults**: Defines the GeoDataPointToEntityQueryResults Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[77]`)
- [ ] `cx:geoChildEntitiesQueryResults` — **GeoChildEntitiesQueryResults**: Defines the GeoChildEntitiesQueryResults Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[78]`)
- [ ] `cx:geoParentEntitiesQueryResults` — **GeoParentEntitiesQueryResults**: Defines the GeoParentEntitiesQueryResults Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[79]`)
- [ ] `cx:binary` — **Xsdbase64Binary**: Defines the Xsdbase64Binary Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[80]`)
- [ ] `cx:clear` — **Clear**: Defines the Clear Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[81]`)
- [ ] `cx:geoCache` — **GeoCache**: Defines the GeoCache Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[82]`)
- [ ] `cx:parentLabelLayout` — **ParentLabelLayout**: Defines the ParentLabelLayout Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[83]`)
- [ ] `cx:regionLabelLayout` — **RegionLabelLayout**: Defines the RegionLabelLayout Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[84]`)
- [ ] `cx:visibility` — **SeriesElementVisibilities**: Defines the SeriesElementVisibilities Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[85]`)
- [ ] `cx:aggregation` — **Aggregation**: Defines the Aggregation Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[86]`)
- [ ] `cx:binning` — **Binning**: Defines the Binning Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[87]`)
- [ ] `cx:geography` — **Geography**: Defines the Geography Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[88]`)
- [ ] `cx:statistics` — **Statistics**: Defines the Statistics Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[89]`)
- [ ] `cx:subtotals` — **Subtotals**: Defines the Subtotals Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[90]`)
- [ ] `cx:extremeValue` — **ExtremeValueColorPosition**: Defines the ExtremeValueColorPosition Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[91]`)
- [ ] `cx:number` — **NumberColorPosition**: Defines the NumberColorPosition Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[92]`)
- [ ] `cx:percent` — **PercentageColorPosition**: Defines the PercentageColorPosition Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[93]`)
- [ ] `cx:min` — **MinValueColorEndPosition**: Defines the MinValueColorEndPosition Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[94]`)
- [ ] `cx:max` — **MaxValueColorEndPosition**: Defines the MaxValueColorEndPosition Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[95]`)
  - _abstract base: `OpenXmlValueColorEndPositionElement`_ — Defines the OpenXmlValueColorEndPositionElement Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[96]`)
- [ ] `cx:mid` — **ValueColorMiddlePosition**: Defines the ValueColorMiddlePosition Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[97]`)
- [ ] `cx:visibility` — **DataLabelVisibilities**: Defines the DataLabelVisibilities Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[98]`)
- [ ] `cx:dataLabel` — **DataLabel**: Defines the DataLabel Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[99]`)
- [ ] `cx:dataLabelHidden` — **DataLabelHidden**: Defines the DataLabelHidden Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[100]`)
- [ ] `cx:valueColors` — **ValueColors**: Defines the ValueColors Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[101]`)
- [ ] `cx:valueColorPositions` — **ValueColorPositions**: Defines the ValueColorPositions Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[102]`)
- [x] `cx:dataPt` — **DataPoint**: Defines the DataPoint Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[103]`)
- [x] `cx:dataLabels` — **DataLabels**: Defines the DataLabels Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[104]`)
- [ ] `cx:dataId` — **DataId**: Defines the DataId Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[105]`)
- [ ] `cx:layoutPr` — **SeriesLayoutProperties**: Defines the SeriesLayoutProperties Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[106]`)
- [x] `cx:axisId` — **AxisId**: Defines the AxisId Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[107]`)
- [ ] `cx:plotSurface` — **PlotSurface**: Defines the PlotSurface Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[108]`)
- [ ] `cx:series` — **Series**: Defines the Series Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[109]`)
- [ ] `cx:plotAreaRegion` — **PlotAreaRegion**: Defines the PlotAreaRegion Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[110]`)
- [ ] `cx:axis` — **Axis**: Defines the Axis Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[111]`)
- [ ] `cx:title` — **ChartTitle**: Defines the ChartTitle Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[112]`)
- [x] `cx:plotArea` — **PlotArea**: Defines the PlotArea Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[113]`)
- [x] `cx:legend` — **Legend**: Defines the Legend Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[114]`)
- [ ] `cx:fmtOvr` — **FormatOverride**: Defines the FormatOverride Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[115]`)
- [ ] `cx:headerFooter` — **HeaderFooter**: Defines the HeaderFooter Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[116]`)
- [ ] `cx:pageMargins` — **PageMargins**: Defines the PageMargins Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[117]`)
- [ ] `cx:pageSetup` — **PageSetup**: Defines the PageSetup Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[118]`)
- [ ] `cx:chartData` — **ChartData**: Defines the ChartData Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[119]`)
- [x] `cx:chart` — **Chart**: Defines the Chart Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[120]`)
- [ ] `cx:clrMapOvr` — **ColorMappingType**: Defines the ColorMappingType Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[121]`)
- [ ] `cx:fmtOvrs` — **FormatOverrides**: Defines the FormatOverrides Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[122]`)
- [ ] `cx:printSettings` — **PrintSettings**: Defines the PrintSettings Class (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[123]`)
- [ ] `cx:idx` — **UnsignedIntegerType**: Index of subtotal data point (`schemas_microsoft_com_office_drawing_2014_chartex.json#Types[124]`)

## 24. Extensions: `cs:` namespace

Namespace: `http://schemas.microsoft.com/office/drawing/2012/chartStyle`
Source: `schemas_microsoft_com_office_drawing_2012_chartStyle.json`

- [ ] `cs:colorStyle` — **ColorStyle**: Defines the ColorStyle Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[0]`)
- [ ] `cs:chartStyle` — **ChartStyle**: Defines the ChartStyle Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[1]`)
- [ ] `cs:variation` — **ColorStyleVariation**: Defines the ColorStyleVariation Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[2]`)
- [ ] `cs:extLst` — **OfficeArtExtensionList**: Defines the OfficeArtExtensionList Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[3]`)
- [ ] `cs:styleClr` — **StyleColor**: Defines the StyleColor Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[4]`)
- [ ] `cs:lnRef` — **LineReference**: Defines the LineReference Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[5]`)
- [ ] `cs:fillRef` — **FillReference**: Defines the FillReference Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[6]`)
- [ ] `cs:effectRef` — **EffectReference**: Defines the EffectReference Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[7]`)
  - _abstract base: `StyleReference`_ — Defines the StyleReference Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[8]`)
- [ ] `cs:lineWidthScale` — **LineWidthScale**: Defines the LineWidthScale Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[9]`)
- [ ] `cs:fontRef` — **FontReference**: Defines the FontReference Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[10]`)
- [ ] `cs:spPr` — **ShapeProperties**: Defines the ShapeProperties Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[11]`)
- [ ] `cs:defRPr` — **TextCharacterPropertiesType**: Defines the TextCharacterPropertiesType Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[12]`)
- [ ] `cs:bodyPr` — **TextBodyProperties**: Defines the TextBodyProperties Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[13]`)
- [ ] `cs:categoryAxis` — **CategoryAxisProperties**: Defines the CategoryAxisProperties Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[14]`)
- [ ] `cs:seriesAxis` — **SeriesAxisProperties**: Defines the SeriesAxisProperties Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[15]`)
- [ ] `cs:valueAxis` — **ValueAxisProperties**: Defines the ValueAxisProperties Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[16]`)
  - _abstract base: `AxisProperties`_ — Defines the AxisProperties Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[17]`)
- [ ] `cs:dataSeries` — **DataSeries**: Defines the DataSeries Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[18]`)
- [x] `cs:dataLabels` — **DataLabels**: Defines the DataLabels Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[19]`)
- [x] `cs:dataTable` — **DataTable**: Defines the DataTable Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[20]`)
- [x] `cs:legend` — **Legend**: Defines the Legend Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[21]`)
- [x] `cs:title` — **Title**: Defines the Title Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[22]`)
- [x] `cs:trendline` — **Trendline**: Defines the Trendline Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[23]`)
- [ ] `cs:view3D` — **View3DProperties**: Defines the View3DProperties Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[24]`)
- [ ] `cs:axisTitle` — **AxisTitle**: Defines the AxisTitle Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[25]`)
- [x] `cs:categoryAxis` — **CategoryAxis**: Defines the CategoryAxis Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[26]`)
- [ ] `cs:chartArea` — **ChartArea**: Defines the ChartArea Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[27]`)
- [ ] `cs:dataLabel` — **DataLabel**: Defines the DataLabel Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[28]`)
- [ ] `cs:dataLabelCallout` — **DataLabelCallout**: Defines the DataLabelCallout Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[29]`)
- [x] `cs:dataPoint` — **DataPoint**: Defines the DataPoint Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[30]`)
- [ ] `cs:dataPoint3D` — **DataPoint3D**: Defines the DataPoint3D Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[31]`)
- [ ] `cs:dataPointLine` — **DataPointLine**: Defines the DataPointLine Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[32]`)
- [ ] `cs:dataPointMarker` — **DataPointMarker**: Defines the DataPointMarker Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[33]`)
- [ ] `cs:dataPointWireframe` — **DataPointWireframe**: Defines the DataPointWireframe Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[34]`)
- [ ] `cs:dataTable` — **DataTableStyle**: Defines the DataTableStyle Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[35]`)
- [ ] `cs:downBar` — **DownBar**: Defines the DownBar Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[36]`)
- [ ] `cs:dropLine` — **DropLine**: Defines the DropLine Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[37]`)
- [ ] `cs:errorBar` — **ErrorBar**: Defines the ErrorBar Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[38]`)
- [ ] `cs:floor` — **Floor**: Defines the Floor Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[39]`)
- [ ] `cs:gridlineMajor` — **GridlineMajor**: Defines the GridlineMajor Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[40]`)
- [ ] `cs:gridlineMinor` — **GridlineMinor**: Defines the GridlineMinor Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[41]`)
- [ ] `cs:hiLoLine` — **HiLoLine**: Defines the HiLoLine Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[42]`)
- [ ] `cs:leaderLine` — **LeaderLine**: Defines the LeaderLine Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[43]`)
- [ ] `cs:legend` — **LegendStyle**: Defines the LegendStyle Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[44]`)
- [x] `cs:plotArea` — **PlotArea**: Defines the PlotArea Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[45]`)
- [ ] `cs:plotArea3D` — **PlotArea3D**: Defines the PlotArea3D Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[46]`)
- [x] `cs:seriesAxis` — **SeriesAxis**: Defines the SeriesAxis Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[47]`)
- [ ] `cs:seriesLine` — **SeriesLine**: Defines the SeriesLine Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[48]`)
- [ ] `cs:title` — **TitleStyle**: Defines the TitleStyle Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[49]`)
- [ ] `cs:trendline` — **TrendlineStyle**: Defines the TrendlineStyle Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[50]`)
- [ ] `cs:trendlineLabel` — **TrendlineLabel**: Defines the TrendlineLabel Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[51]`)
- [ ] `cs:upBar` — **UpBar**: Defines the UpBar Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[52]`)
- [x] `cs:valueAxis` — **ValueAxis**: Defines the ValueAxis Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[53]`)
- [ ] `cs:wall` — **Wall**: Defines the Wall Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[54]`)
  - _abstract base: `StyleEntry`_ — Defines the StyleEntry Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[55]`)
- [ ] `cs:dataPointMarkerLayout` — **MarkerLayoutProperties**: Defines the MarkerLayoutProperties Class (`schemas_microsoft_com_office_drawing_2012_chartStyle.json#Types[56]`)

## 25. Extensions: `c16r3:` namespace

Namespace: `http://schemas.microsoft.com/office/drawing/2017/03/chart`
Source: `schemas_microsoft_com_office_drawing_2017_03_chart.json`

- [ ] `c16r3:dataDisplayOptions16` — **DataDisplayOptions16**: Defines the DataDisplayOptions16 Class (`schemas_microsoft_com_office_drawing_2017_03_chart.json#Types[0]`)
- [ ] `c16r3:dispNaAsBlank` — **BooleanFalse**: Defines the BooleanFalse Class (`schemas_microsoft_com_office_drawing_2017_03_chart.json#Types[1]`)

## 26. Enumerations

Enum types referenced in chart/drawing attribute definitions.
Values are defined in the SDK's generated code, not in the schema JSONs.

- [ ] `AxisPositionValues`
- [ ] `AxisUnit`
- [ ] `BarDirectionValues`
- [ ] `BarGroupingValues`
- [ ] `BevelPresetValues`
- [ ] `BlackWhiteModeValues`
- [ ] `BlendModeValues`
- [ ] `BlipCompressionValues`
- [ ] `Boolean`
- [ ] `BooleanStyleValues`
- [ ] `BuiltInUnitValues`
- [ ] `ChartBuildStepValues`
- [ ] `ColorSchemeIndexValues`
- [ ] `CompoundLineValues`
- [ ] `CrossBetweenValues`
- [ ] `CrossesValues`
- [ ] `DataLabelPos`
- [ ] `DataLabelPositionValues`
- [ ] `DataLabelsPosition`
- [ ] `DiagramBuildStepValues`
- [ ] `DisplayBlanksAsValues`
- [ ] `EditAsValues`
- [ ] `EffectContainerValues`
- [ ] `EntityTypeEnum`
- [ ] `ErrorBarDirectionValues`
- [ ] `ErrorBarValues`
- [ ] `ErrorValues`
- [ ] `FontCollectionIndexValues`
- [ ] `FormulaDirection`
- [ ] `GeoMappingLevel`
- [ ] `GeoProjectionType`
- [ ] `GroupingValues`
- [ ] `IntervalClosedSide`
- [ ] `LabelAlignmentValues`
- [ ] `LayoutModeValues`
- [ ] `LayoutTargetValues`
- [ ] `LegendPosition`
- [ ] `LegendPositionValues`
- [ ] `LightRigDirectionValues`
- [ ] `LightRigValues`
- [ ] `LineCapValues`
- [ ] `LineEndLengthValues`
- [ ] `LineEndValues`
- [ ] `LineEndWidthValues`
- [ ] `MarkerStyle`
- [ ] `MarkerStyleValues`
- [ ] `NumericDimensionType`
- [ ] `OfPieValues`
- [ ] `OrientationValues`
- [ ] `PageOrientation`
- [ ] `PageSetupOrientationValues`
- [ ] `ParentLabelLayoutVal`
- [ ] `PathFillModeValues`
- [ ] `PathShadeValues`
- [ ] `PenAlignmentValues`
- [ ] `PictureFormatValues`
- [ ] `PosAlign`
- [ ] `PresetCameraValues`
- [ ] `PresetColorValues`
- [ ] `PresetLineDashValues`
- [ ] `PresetMaterialTypeValues`
- [ ] `PresetPatternValues`
- [ ] `PresetShadowValues`
- [ ] `QuartileMethod`
- [ ] `RadarStyleValues`
- [ ] `RectangleAlignmentValues`
- [ ] `RegionLabelLayoutEnum`
- [ ] `ScatterStyleValues`
- [ ] `SchemeColorValues`
- [ ] `SeriesLayout`
- [ ] `ShapeTypeValues`
- [ ] `ShapeValues`
- [ ] `SidePos`
- [ ] `SizeRepresentsValues`
- [ ] `SplitValues`
- [ ] `StringDimensionType`
- [ ] `SystemColorValues`
- [ ] `TextAlignmentTypeValues`
- [ ] `TextAnchoringTypeValues`
- [ ] `TextAutoNumberSchemeValues`
- [ ] `TextCapsValues`
- [ ] `TextFontAlignmentValues`
- [ ] `TextHorizontalOverflowValues`
- [ ] `TextShapeValues`
- [ ] `TextStrikeValues`
- [ ] `TextTabAlignmentValues`
- [ ] `TextUnderlineValues`
- [ ] `TextVerticalOverflowValues`
- [ ] `TextVerticalValues`
- [ ] `TextWrappingValues`
- [ ] `TickLabelPositionNinch`
- [ ] `TickLabelPositionValues`
- [ ] `TickMarkNinch`
- [ ] `TickMarkValues`
- [ ] `TickMarksType`
- [ ] `TileFlipValues`
- [ ] `TimeUnitValues`
- [ ] `TitlePosition`
- [ ] `TrendlineValues`

