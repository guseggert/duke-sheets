use std::cell::RefCell;
use std::io::Cursor;
use std::rc::Rc;
use std::sync::Arc;

use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use wasm_bindgen::JsCast;

use duke_sheets::{CalculationOptions, FormulaValue, WorkbookCalculationExt};
use duke_sheets_core::{
    CellAddress, CellError, CellRange, CellValue as CoreCellValue, Workbook as CoreWorkbook,
};
use duke_sheets_xlsb::XlsbWriter;
use duke_sheets_xlsx::XlsxWriter;

fn parse_xlsx_profile(
    profile: Option<&str>,
    key_bits: Option<u32>,
    spin_count: Option<u32>,
) -> std::result::Result<duke_sheets_xlsx::EncryptionProfile, String> {
    use duke_sheets_xlsx::EncryptionProfile;
    let normalized = profile.map(|p| p.to_lowercase());
    Ok(match normalized.as_deref() {
        None | Some("default") | Some("agile") | Some("ooxml-agile") => EncryptionProfile::Agile {
            key_bits: key_bits.unwrap_or(256),
            spin_count: spin_count.unwrap_or(100_000),
        },
        Some("standard") | Some("ooxml-standard") => EncryptionProfile::Standard {
            key_bits: key_bits.unwrap_or(128),
        },
        Some(other) => return Err(format!("unknown XLSX encryption profile: {other:?}")),
    })
}

fn parse_xls_variant(
    profile: Option<&str>,
    key_bits: Option<u32>,
) -> std::result::Result<duke_sheets_crypto::xls::XlsEncryptionVariant, String> {
    use duke_sheets_crypto::xls::XlsEncryptionVariant;
    let normalized = profile.map(|p| p.to_lowercase());
    Ok(match normalized.as_deref() {
        None | Some("default") | Some("rc4-cryptoapi") | Some("xls-rc4-cryptoapi") => {
            XlsEncryptionVariant::Rc4CryptoApi {
                key_bits: key_bits.unwrap_or(128),
            }
        }
        Some("rc4-legacy") | Some("xls-rc4-legacy") => XlsEncryptionVariant::Rc4Legacy,
        Some("xor") | Some("xls-xor") => XlsEncryptionVariant::Xor,
        Some(other) => return Err(format!("unknown XLS encryption profile: {other:?}")),
    })
}

mod types;
pub(crate) mod drawings;
mod workbook_read;
mod worksheet_read;

pub use types::*;

#[wasm_bindgen(typescript_custom_section)]
const ROW_ITERATOR_TYPES: &str = r#"
declare global {
  interface SymbolConstructor {
    readonly dispose: unique symbol;
  }
}

export interface JsRowCell {
  col: number;
  value: string;
  style?: any;
  mergeSpan?: JsMergeSpan;
  isMergedSecondary?: boolean;
  hyperlink?: any;
  comment?: any;
  formula?: string;
  image?: any;
}

export interface JsRow {
  index: number;
  cells: Array<JsRowCell>;
}

export interface JsMergeSpan {
  rowSpan: number;
  colSpan: number;
}

export interface JsRowsOptions {
  useFormattedValues?: boolean;
  useCalculatedValues?: boolean;
  includeStyles?: boolean;
  includeMergeInfo?: boolean;
  includeHyperlinks?: boolean;
  includeComments?: boolean;
  includeFormulas?: boolean;
  includeImages?: boolean;
  skipEmptyValues?: boolean;
  skipBlankValues?: boolean;
}

export interface ColorInput {
  colorType?: "auto" | "rgb" | "argb" | "theme" | "indexed";
  hex?: string;
  r?: number;
  g?: number;
  b?: number;
  a?: number;
  themeIndex?: number;
  tint?: number;
  paletteIndex?: number;
}

export interface FontStyleInput {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
  strikethrough?: boolean;
  color?: ColorInput;
  verticalAlign?: "baseline" | "superscript" | "subscript";
  family?: number;
  charset?: number;
  scheme?: string;
}

export interface GradientStopInput {
  position: number;
  color: ColorInput;
}

export interface FillStyleInput {
  fillType?: "none" | "solid" | "pattern" | "gradient";
  color?: ColorInput;
  pattern?: string;
  foreground?: ColorInput;
  background?: ColorInput;
  gradientType?: "linear" | "path";
  angle?: number;
  stops?: GradientStopInput[];
}

export interface BorderEdgeInput {
  style?: "none" | "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
  color?: ColorInput;
}

export interface BorderStyleInput {
  left?: BorderEdgeInput;
  right?: BorderEdgeInput;
  top?: BorderEdgeInput;
  bottom?: BorderEdgeInput;
  diagonal?: BorderEdgeInput;
  diagonalDirection?: "none" | "down" | "up" | "both";
}

export interface AlignmentInput {
  horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  rotation?: number;
  readingOrder?: "contextDependent" | "leftToRight" | "rightToLeft";
}

export interface NumberFormatInput {
  formatType?: "general" | "builtin" | "custom";
  id?: number;
  formatString?: string;
}

export interface CellProtectionInput {
  locked?: boolean;
  hidden?: boolean;
}

export interface StyleInput {
  font?: FontStyleInput;
  fill?: FillStyleInput;
  border?: BorderStyleInput;
  alignment?: AlignmentInput;
  numberFormat?: NumberFormatInput;
  protection?: CellProtectionInput;
}

export class RowIterator implements IterableIterator<JsRow> {
  constructor(ws: Worksheet, opts?: JsRowsOptions, maxRow?: number);
  [Symbol.iterator](): RowIterator;
  next(): IteratorResult<JsRow>;
}

export interface DrawingMeta {
  name?: string;
  hidden: boolean;
  locked: boolean;
  printable: boolean;
  altText?: string;
  title?: string;
}

export interface DrawingMetaInput {
  name?: string;
  hidden?: boolean;
  locked?: boolean;
  printable?: boolean;
  altText?: string;
  title?: string;
}

export interface DrawingCellMarker {
  col: number;
  row: number;
  colOffsetEmu?: number;
  rowOffsetEmu?: number;
}

export type DrawingAnchor =
  | { type: "twoCell"; from: DrawingCellMarker; to: DrawingCellMarker; editAs?: "twoCell" | "oneCell" | "absolute" }
  | { type: "oneCell"; from: DrawingCellMarker; widthEmu: number; heightEmu: number }
  | { type: "absolute"; xEmu: number; yEmu: number; widthEmu: number; heightEmu: number };

export interface DrawingChildTransform {
  xEmu?: number;
  yEmu?: number;
  cxEmu?: number;
  cyEmu?: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export interface DrawingGroupTransform extends DrawingChildTransform {
  childXEmu?: number;
  childYEmu?: number;
  childCxEmu?: number;
  childCyEmu?: number;
}

export type DrawingPlacement =
  | { anchor: DrawingAnchor; transform?: never }
  | { anchor?: never; transform: DrawingChildTransform };

export type DrawingColor =
  | { colorType: "auto" }
  | { colorType: "rgb"; r: number; g: number; b: number }
  | { colorType: "argb"; a: number; r: number; g: number; b: number }
  | { colorType: "theme"; index: number; tint: number }
  | { colorType: "indexed"; index: number };

export interface DrawingRunFont {
  bold?: boolean;
  italic?: boolean;
  size?: number;
  color?: DrawingColor;
  name?: string;
  underline?: "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
  strikethrough?: boolean;
  verticalAlign?: "baseline" | "superscript" | "subscript";
  family?: number;
  charset?: number;
  scheme?: string;
}

export interface DrawingText {
  runs: Array<{ text: string; font?: DrawingRunFont }>;
  horizontalAlignment?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  verticalAlignment?: "top" | "center" | "bottom" | "justify" | "distributed";
}

export type FormControlKind =
  | { kind: "button"; caption: DrawingText }
  | { kind: "checkbox"; caption: DrawingText; state: "unchecked" | "checked" | "mixed"; cellLink?: string; no3D: boolean }
  /**
   * `firstInGroup` reports whether this radio heads its group; writers
   * recompute it from group-box containment, so it is read-side
   * information. `"mixed"` never validates on write but can surface
   * when reading hostile files.
   */
  | { kind: "optionButton"; caption: DrawingText; state: "unchecked" | "checked" | "mixed"; cellLink?: string; firstInGroup: boolean; no3D: boolean }
  | { kind: "label"; caption: DrawingText }
  | { kind: "groupBox"; caption: DrawingText; no3D: boolean }
  | { kind: "listBox"; inputRange?: string; cellLink?: string; selection: "single" | "multi" | "extend"; selected: number[]; no3D: boolean }
  | { kind: "dropdown"; inputRange?: string; cellLink?: string; selected?: number; lines: number; no3D: boolean }
  | { kind: "scrollbar"; value: number; min: number; max: number; increment: number; page: number; horizontal: boolean; cellLink?: string }
  | { kind: "spinner"; value: number; min: number; max: number; increment: number; cellLink?: string }
  /** Unsupported legacy control, preserved for passthrough. */
  | { kind: "unknown"; objectType: string; legacyObjectType?: number; caption: DrawingText };

export type FormControlKindInput =
  | { kind: "button"; caption: DrawingText }
  | { kind: "checkbox"; caption: DrawingText; state: "unchecked" | "checked" | "mixed"; cellLink?: string; no3D?: boolean }
  /** `firstInGroup` is ignored on input; writers recompute it from group-box containment. */
  | { kind: "optionButton"; caption: DrawingText; state: "unchecked" | "checked"; cellLink?: string; firstInGroup?: boolean; no3D?: boolean }
  | { kind: "label"; caption: DrawingText }
  | { kind: "groupBox"; caption: DrawingText; no3D?: boolean }
  | { kind: "listBox"; inputRange?: string; cellLink?: string; selection: "single" | "multi" | "extend"; selected?: number[]; no3D?: boolean }
  | { kind: "dropdown"; inputRange?: string; cellLink?: string; selected?: number; lines: number; no3D?: boolean }
  | { kind: "scrollbar"; value: number; min: number; max: number; increment: number; page: number; horizontal?: boolean; cellLink?: string }
  | { kind: "spinner"; value: number; min: number; max: number; increment: number; cellLink?: string }
  | { kind: "unknown"; objectType: string; legacyObjectType?: number; caption?: DrawingText };

export interface FormControlPayload {
  kind: FormControlKind;
  macroName?: string;
  /**
   * Unmodeled VML ClientData child fragments preserved on any control
   * kind; opaque internal passthrough echoed back unchanged on write.
   */
  rawClientData: number[][];
  /**
   * Unmodeled XLSX formControlPr attributes preserved on any control
   * kind; opaque internal passthrough echoed back unchanged on write.
   */
  rawProperties: Array<[string, string]>;
  /** Original BIFF OBJ body for XLS passthrough of unknown controls. */
  rawObj?: number[];
}

export interface FormControlInputPayload {
  kind: FormControlKindInput;
  macroName?: string;
  /** Echo back the raw* fields unchanged when rewriting a control read from a file. */
  rawClientData?: Array<Uint8Array | number[]>;
  rawProperties?: Array<[string, string]>;
  rawObj?: Uint8Array | number[];
}

export interface DrawingImage {
  format: "png" | "jpeg" | "gif" | "bmp" | "emf" | "wmf" | "tiff" | "svg";
  mediaPath: string;
  svgMediaPath?: string;
  widthEmu: number;
  heightEmu: number;
  rotation?: number;
  flipH: boolean;
  flipV: boolean;
}

export interface DrawingImageInput extends Partial<Omit<DrawingImage, "format" | "widthEmu" | "heightEmu">> {
  format: DrawingImage["format"];
  widthEmu: number;
  heightEmu: number;
  data: Uint8Array | number[];
  svgData?: Uint8Array | number[];
}

export type DrawingShapeFill =
  | { kind: "none" }
  | { kind: "solid"; color: DrawingColor };

export interface DrawingShape {
  geometry: string;
  fill: DrawingShapeFill;
  line: { color?: DrawingColor; widthEmu?: number; dashStyle?: string; noFill: boolean };
  text?: DrawingText;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
}

export interface DrawingShapeInput extends Partial<DrawingShape> {
  geometry?: string;
}

export interface DrawingComment {
  row: number;
  col: number;
  author: string;
  /** Plain text (runs concatenated). */
  text: string;
  /**
   * Rich runs; present on output when any run is formatted, and wins
   * over `text` on input when supplied.
   */
  richText?: DrawingText;
}

export type ChartDataReference =
  | { refType: "formula"; formula: string }
  | { refType: "numbers"; numbers: number[] }
  | { refType: "strings"; strings: string[] };

export interface ChartShapeProperties {
  solidFillHex?: string;
  noFill: boolean;
  lineWidth?: number;
  lineColorHex?: string;
  lineNoFill: boolean;
  lineDashStyle?: string;
}

export interface ChartNumberFormat {
  formatCode: string;
  sourceLinked?: boolean;
}

export interface ChartDataLabels {
  showLegendKey?: boolean;
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
  position?: string;
  numberFormat?: ChartNumberFormat;
  showLeaderLines?: boolean;
}

export interface ChartDataSeries {
  name?: string;
  values: ChartDataReference;
  categories?: ChartDataReference;
  dataLabels?: ChartDataLabels;
  trendline?: { trendlineType: string; name?: string; order?: number; period?: number; forward?: number; backward?: number; intercept?: number; displayRSquared?: boolean; displayEquation?: boolean };
  errorBars?: { direction: string; barType: string; valueType: string; value?: number; noEndCap?: boolean };
  marker?: { symbol?: string; size?: number };
  dataPoints: Array<{ index: number; marker?: { symbol?: string; size?: number }; explosion?: number; shapeProperties?: ChartShapeProperties }>;
  smooth?: boolean;
  explosion?: number;
  invertIfNegative?: boolean;
  shapeProperties?: ChartShapeProperties;
}

export interface ChartAxis {
  title?: string;
  minimum?: number;
  maximum?: number;
  majorUnit?: number;
  minorUnit?: number;
  position: "bottom" | "top" | "left" | "right";
  numberFormat?: ChartNumberFormat;
  majorGridlines: boolean;
  minorGridlines: boolean;
  majorGridlinesShapeProperties?: ChartShapeProperties;
  minorGridlinesShapeProperties?: ChartShapeProperties;
  majorTickMark?: string;
  minorTickMark?: string;
  labelPosition?: string;
  delete?: boolean;
  crosses?: string;
  crossBetween?: string;
  shapeProperties?: ChartShapeProperties;
}

export interface Chart {
  chartType: string;
  title?: string;
  series: ChartDataSeries[];
  categoryAxis?: ChartAxis;
  valueAxis?: ChartAxis;
  legend?: { position: string; overlay: boolean };
  dataLabels?: ChartDataLabels;
  view3D?: { rotateX?: number; rotateY?: number; depthPercent?: number; heightPercent?: number; perspective?: number; rightAngleAxes?: boolean };
  dataTable?: { showHorizontalBorder?: boolean; showVerticalBorder?: boolean; showOutline?: boolean; showKeys?: boolean };
  displayBlanksAs?: "gap" | "span" | "zero";
  plotVisibleOnly?: boolean;
  layout?: { manualLayout?: { x?: number; y?: number; width?: number; height?: number } };
  shapeProperties?: ChartShapeProperties;
  is3D: boolean;
  varyColors?: boolean;
  gapWidth?: number;
  overlap?: number;
  firstSliceAngle?: number;
  holeSize?: number;
  bubbleScale?: number;
  showNegativeBubbles?: boolean;
  autoTitleDeleted?: boolean;
  roundedCorners?: boolean;
  showDlblsOverMax?: boolean;
  wireframe?: boolean;
  radarStyle?: string;
  typeGroups: Array<{ chartType: string; is3D: boolean; series: ChartDataSeries[]; dataLabels?: ChartDataLabels; varyColors?: boolean; gapWidth?: number; overlap?: number; firstSliceAngle?: number; holeSize?: number; bubbleScale?: number; showNegativeBubbles?: boolean; radarStyle?: string; wireframe?: boolean; axisIds: number[] }>;
  axes: Array<{ id: number; crossId: number; axis: ChartAxis }>;
}

export interface ChartSeriesInput {
  name?: string;
  values: ChartDataReference;
  categories?: ChartDataReference;
}

/** Exactly the chart fields accepted when authoring; all other Chart fields are read-only. */
export interface ChartInput {
  chartType: string;
  title?: string;
  series?: ChartSeriesInput[];
  is3D?: boolean;
  varyColors?: boolean;
  gapWidth?: number;
  overlap?: number;
}

export interface ChartExTitle {
  text?: string;
  position?: string;
  align?: string;
  overlay?: boolean;
  offset?: { top?: number; left?: number };
  shapeProperties?: ChartShapeProperties;
}

export interface ChartExGridlines {
  shapeProperties?: ChartShapeProperties;
}

export interface ChartExDataLabelOverride {
  idx: number;
  position?: string;
  visibilitySeriesName?: boolean;
  visibilityCategoryName?: boolean;
  visibilityValue?: boolean;
  numberFormat?: ChartNumberFormat;
  separator?: string;
  shapeProperties?: ChartShapeProperties;
}

export interface ChartExColorPosition {
  positionType: 'extremeValue' | 'number' | 'percent';
  value?: number;
}

export interface ChartExSeriesLayoutProperties {
  parentLabelLayout?: string;
  regionLabelLayout?: string;
  visibility?: { connectorLines?: boolean; meanLine?: boolean; meanMarker?: boolean; nonoutliers?: boolean; outliers?: boolean };
  aggregation: boolean;
  binning?: { intervalClosed?: string; underflow?: string; overflow?: string; binSize?: number; binCount?: number };
  geography?: { projectionType?: string; viewedRegionType?: string; cultureLanguage?: string; cultureRegion?: string; attribution?: string };
  statistics?: { quartileMethod?: string };
  /** Absent when the cx:subtotals element is absent; [] when it is present but empty. */
  subtotals?: number[];
}

export interface ChartExSeries {
  layout: string;
  dataId: number;
  uniqueId?: string;
  hidden?: boolean;
  ownerIdx?: number;
  formatIdx?: number;
  text?: { formula?: string; value?: string };
  dataLabels?: { position?: string; visibilitySeriesName?: boolean; visibilityCategoryName?: boolean; visibilityValue?: boolean; numberFormat?: ChartNumberFormat; separator?: string; shapeProperties?: ChartShapeProperties; overrides: ChartExDataLabelOverride[]; hiddenLabels: number[] };
  dataPoints: Array<{ idx: number; shapeProperties?: ChartShapeProperties }>;
  layoutProperties?: ChartExSeriesLayoutProperties;
  axisIds: number[];
  valueColors: boolean;
  valueColorPositions?: { count?: number; min?: ChartExColorPosition; mid?: ChartExColorPosition; max?: ChartExColorPosition };
  shapeProperties?: ChartShapeProperties;
}

export interface ChartExAxis {
  id: number;
  hidden?: boolean;
  scaling: { scalingType: string; gapWidth?: number; min?: number; max?: number; majorUnit?: number; minorUnit?: number };
  title?: { text?: string; shapeProperties?: ChartShapeProperties };
  units?: { unit?: string };
  majorGridlines?: ChartExGridlines;
  minorGridlines?: ChartExGridlines;
  majorTickMarks?: string;
  minorTickMarks?: string;
  tickLabels: boolean;
  numberFormat?: ChartNumberFormat;
  shapeProperties?: ChartShapeProperties;
}

export interface ChartEx {
  layout: string;
  version?: string;
  featureList?: string;
  fallbackImg?: string;
  title?: ChartExTitle;
  data: Array<{ id: number; dimensions: Array<{ dimType: string; formula?: string; nfFormula?: string }> }>;
  plotArea: { plotSurface?: ChartShapeProperties; series: ChartExSeries[]; axes: ChartExAxis[]; shapeProperties?: ChartShapeProperties };
  legend?: { position?: string; align?: string; overlay?: boolean; offset?: { top?: number; left?: number }; shapeProperties?: ChartShapeProperties };
  shapeProperties?: ChartShapeProperties;
  formatOverrides: Array<{ idx: number; shapeProperties?: ChartShapeProperties }>;
  externalDataRelId?: string;
  externalDataAutoUpdate?: boolean;
}

export interface ChartExTitleInput {
  text?: string;
  position?: string;
  align?: string;
  overlay?: boolean;
}

/** Exactly the ChartEx fields accepted when authoring; all other ChartEx fields are read-only. */
export interface ChartExInput {
  layout: string;
  version?: string;
  featureList?: string;
  fallbackImg?: string;
  title?: ChartExTitleInput;
}

export interface RawDrawingMetadata {
  byteLength: number;
  relationships: Array<{ id: string; relType: string; target: string; external: boolean; hasPart: boolean }>;
}

/**
 * Resolved on-sheet placement in EMU: the anchor rectangle for
 * top-level drawings, the group-mapped (rotation/flip aware)
 * rectangle for group children.
 */
export type RectEmu = { xEmu: number; yEmu: number; widthEmu: number; heightEmu: number };

type DrawingNode = DrawingMeta & DrawingPlacement & { drawingPath: number[]; absoluteRectEmu: RectEmu };

export type ImageDrawing = DrawingNode & { kind: "image"; image: DrawingImage };
export type ChartDrawing = DrawingNode & { kind: "chart"; chart: Chart };
export type ChartExDrawing = DrawingNode & { kind: "chartEx"; chartEx: ChartEx };
export type FormControlDrawing = DrawingNode & { kind: "formControl"; formControl: FormControlPayload };
export type CommentDrawing = DrawingNode & { kind: "comment"; comment: DrawingComment };
export type ShapeDrawing = DrawingNode & { kind: "shape"; shape: DrawingShape };
export type GroupDrawing = DrawingNode & { kind: "group"; group: { groupTransform: DrawingGroupTransform; children: Drawing[] } };
export type RawDrawing = DrawingNode & { kind: "raw"; raw: RawDrawingMetadata };

export type Drawing = ImageDrawing | ChartDrawing | ChartExDrawing | FormControlDrawing | CommentDrawing | ShapeDrawing | GroupDrawing | RawDrawing;

type DrawingInputPlacement =
  | { anchor: DrawingAnchor; transform?: never }
  | { anchor?: never; transform: DrawingChildTransform };

export type DrawingInput = DrawingMetaInput & DrawingInputPlacement & (
  | { kind: "image"; image: DrawingImageInput }
  | { kind: "chart"; chart: ChartInput }
  | { kind: "chartEx"; chartEx: ChartExInput }
  | { kind: "formControl"; formControl: FormControlInputPayload }
  | { kind: "comment"; comment: DrawingComment }
  | { kind: "shape"; shape: DrawingShapeInput }
  | { kind: "group"; group: { groupTransform?: DrawingGroupTransform; children?: DrawingInput[] } }
);

export interface FormControlInteractionResult {
  controlsChanged: number;
  linkedCellsChanged: number;
}

export interface Worksheet {
  iterateRows(opts?: JsRowsOptions): RowIterator;
  setCellStyle(address: string, style: StyleInput): void;
  setCellStyleAt(row: number, col: number, style: StyleInput): void;
  setRangeStyle(range: string, style: StyleInput): void;
  readonly drawings: Drawing[];
  readonly formControls: FormControlDrawing[];
  readonly formControlCount: number;
  readonly images: ImageDrawing[];
  readonly imageCount: number;
  readonly charts: ChartDrawing[];
  readonly chartCount: number;
  readonly chartsEx: ChartExDrawing[];
  readonly chartExCount: number;
  addDrawing(drawing: DrawingInput & { anchor: DrawingAnchor }): number;
  /** Drawing paths are positional; mutating the list invalidates previously returned paths. */
  insertDrawing(index: number, drawing: DrawingInput & { anchor: DrawingAnchor }): void;
  /** Drawing paths are positional; mutating the list invalidates previously returned paths. */
  setDrawing(path: number[], drawing: DrawingInput): void;
  /** Drawing paths are positional; mutating the list invalidates previously returned paths. */
  removeDrawing(path: number[]): void;
  /** Drawing paths are positional; mutating the list invalidates previously returned paths. */
  moveDrawing(from: number, to: number): void;
  /** Paths are positional; mutating the drawing list invalidates previously returned paths. */
  drawingImageData(path: number[]): Uint8Array;
  /** Paths are positional; mutating the drawing list invalidates previously returned paths. */
  drawingSvgData(path: number[]): Uint8Array | undefined;
  setFormControlCheckState(path: number[], state: "unchecked" | "checked" | "mixed"): FormControlInteractionResult;
}

export interface Workbook {
  /**
   * Resolve a drawing color to display RGB ("RRGGBB" hex) against
   * this workbook's theme palette; `auto` resolves to undefined.
   */
  resolveColor(color: DrawingColor): string | undefined;
}
"#;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct JsCalculationOptions {
    iterative: Option<bool>,
    max_iterations: Option<u32>,
    max_change: Option<f64>,
    force_full_calculation: Option<bool>,
    calculate_volatile: Option<bool>,
    sheets: Option<Vec<usize>>,
    max_threads: Option<usize>,
}

/// Wrapper to make `js_sys::Function` implement Send + Sync.
/// SAFETY: WASM is single-threaded, so Send + Sync are no-ops.
struct SendSyncFunction(js_sys::Function);
unsafe impl Send for SendSyncFunction {}
unsafe impl Sync for SendSyncFunction {}

impl SendSyncFunction {
    fn call1(&self, this: &JsValue, arg: &JsValue) -> Result<JsValue, JsValue> {
        self.0.call1(this, arg)
    }

    fn call2(&self, this: &JsValue, arg1: &JsValue, arg2: &JsValue) -> Result<JsValue, JsValue> {
        self.0.call2(this, arg1, arg2)
    }

    fn call3(
        &self,
        this: &JsValue,
        arg1: &JsValue,
        arg2: &JsValue,
        arg3: &JsValue,
    ) -> Result<JsValue, JsValue> {
        self.0.call3(this, arg1, arg2, arg3)
    }
}

pub(crate) fn to_js_error(e: impl std::fmt::Display) -> JsError {
    JsError::new(&e.to_string())
}

pub(crate) fn to_js_value<T: Serialize>(value: &T) -> Result<JsValue, JsError> {
    // Emit plain objects rather than ES2015 Maps for map-shaped output
    // (serde flattens structs, e.g. drawing nodes, through serialize_map),
    // matching the declared TypeScript types.
    value
        .serialize(&serde_wasm_bindgen::Serializer::new().serialize_maps_as_objects(true))
        .map_err(to_js_error)
}

fn cell_error_to_string(e: &CellError) -> &'static str {
    match e {
        CellError::Div0 => "#DIV/0!",
        CellError::Na => "#N/A",
        CellError::Name => "#NAME?",
        CellError::Null => "#NULL!",
        CellError::Num => "#NUM!",
        CellError::Ref => "#REF!",
        CellError::Value => "#VALUE!",
        CellError::GettingData => "#GETTING_DATA",
        CellError::Spill => "#SPILL!",
        CellError::Calc => "#CALC!",
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmCalculationStats {
    formula_count: u32,
    cells_calculated: u32,
    errors: u32,
    circular_references: u32,
    volatile_cells: u32,
    converged: bool,
    iterations: u32,
}

impl From<&duke_sheets::CalculationStats> for WasmCalculationStats {
    fn from(stats: &duke_sheets::CalculationStats) -> Self {
        Self {
            formula_count: stats.formula_count as u32,
            cells_calculated: stats.cells_calculated as u32,
            errors: stats.errors as u32,
            circular_references: stats.circular_references as u32,
            volatile_cells: stats.volatile_cells as u32,
            converged: stats.converged,
            iterations: stats.iterations as u32,
        }
    }
}

#[wasm_bindgen]
pub struct CellValue {
    inner: CoreCellValue,
}

#[wasm_bindgen]
impl CellValue {
    #[wasm_bindgen(getter)]
    pub fn is_empty(&self) -> bool {
        matches!(self.inner, CoreCellValue::Empty)
    }

    #[wasm_bindgen(getter)]
    pub fn is_number(&self) -> bool {
        matches!(self.inner, CoreCellValue::Number(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_text(&self) -> bool {
        matches!(self.inner, CoreCellValue::String(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_boolean(&self) -> bool {
        matches!(self.inner, CoreCellValue::Boolean(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_error(&self) -> bool {
        matches!(self.inner, CoreCellValue::Error(_))
    }

    #[wasm_bindgen(js_name = asNumber)]
    pub fn as_number(&self) -> Option<f64> {
        match &self.inner {
            CoreCellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asText)]
    pub fn as_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::String(s) => Some(s.to_string()),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asBoolean)]
    pub fn as_boolean(&self) -> Option<bool> {
        match &self.inner {
            CoreCellValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asError)]
    pub fn as_error(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::Error(e) => Some(cell_error_to_string(e).to_string()),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = toJs)]
    pub fn to_js(&self) -> JsValue {
        match &self.inner {
            CoreCellValue::Empty => JsValue::NULL,
            CoreCellValue::Number(n) => JsValue::from_f64(*n),
            CoreCellValue::String(s) => JsValue::from_str(&s.to_string()),
            CoreCellValue::Boolean(b) => JsValue::from_bool(*b),
            CoreCellValue::Error(e) => JsValue::from_str(cell_error_to_string(e)),
            CoreCellValue::RichText(runs) => {
                let text: String = runs.iter().map(|r| r.text.as_str()).collect();
                JsValue::from_str(&text)
            }
            CoreCellValue::SpillTarget { .. } => JsValue::NULL,
        }
    }

    #[wasm_bindgen(js_name = toString)]
    pub fn to_string_js(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => String::new(),
            CoreCellValue::Number(n) => n.to_string(),
            CoreCellValue::String(s) => s.to_string(),
            CoreCellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CoreCellValue::Error(e) => cell_error_to_string(e).to_string(),
            CoreCellValue::RichText(runs) => runs.iter().map(|r| r.text.as_str()).collect(),
            CoreCellValue::SpillTarget { .. } => String::new(),
        }
    }
}

#[wasm_bindgen]
pub struct Worksheet {
    workbook: Rc<RefCell<CoreWorkbook>>,
    sheet_index: usize,
}

#[wasm_bindgen]
impl Worksheet {
    #[wasm_bindgen(getter)]
    pub fn name(&self) -> Result<String, JsError> {
        let wb = self.workbook.borrow();
        wb.worksheet(self.sheet_index)
            .map(|ws| ws.name().to_string())
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))
    }

    #[wasm_bindgen(js_name = setCell)]
    pub fn set_cell(&self, address: &str, value: JsValue) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let cell_value = js_to_cell_value(value)?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        ws.set_cell_value_at(addr.row, addr.col, cell_value)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setFormula)]
    pub fn set_formula(&self, address: &str, formula: &str) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_cell_formula(address, formula).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setCellStyle)]
    pub fn set_cell_style(&self, address: &str, style: JsValue) -> Result<(), JsError> {
        let patch: WasmStylePatch = serde_wasm_bindgen::from_value(style).map_err(to_js_error)?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        let mut core_style = ws
            .cell_style_at(addr.row, addr.col)
            .cloned()
            .unwrap_or_default();
        patch
            .apply_to_core_style(&mut core_style)
            .map_err(to_js_error)?;
        ws.set_cell_style_at(addr.row, addr.col, &core_style)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setCellStyleAt)]
    pub fn set_cell_style_at(&self, row: u32, col: u32, style: JsValue) -> Result<(), JsError> {
        let patch: WasmStylePatch = serde_wasm_bindgen::from_value(style).map_err(to_js_error)?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let mut core_style = ws
            .cell_style_at(row, col as u16)
            .cloned()
            .unwrap_or_default();
        patch
            .apply_to_core_style(&mut core_style)
            .map_err(to_js_error)?;
        ws.set_cell_style_at(row, col as u16, &core_style)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setRangeStyle)]
    pub fn set_range_style(&self, range_str: &str, style: JsValue) -> Result<(), JsError> {
        let patch: WasmStylePatch = serde_wasm_bindgen::from_value(style).map_err(to_js_error)?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        for addr in range.cells() {
            let mut core_style = ws
                .cell_style_at(addr.row, addr.col)
                .cloned()
                .unwrap_or_default();
            patch
                .apply_to_core_style(&mut core_style)
                .map_err(to_js_error)?;
            ws.set_cell_style_at(addr.row, addr.col, &core_style)
                .map_err(to_js_error)?;
        }
        Ok(())
    }

    #[wasm_bindgen(js_name = setProtection)]
    pub fn set_protection(&self, protection: JsValue) -> Result<(), JsError> {
        let protection = if protection.is_null() || protection.is_undefined() {
            None
        } else {
            let input: WasmSheetProtectionInput =
                serde_wasm_bindgen::from_value(protection).map_err(to_js_error)?;
            Some(input.into_core().map_err(to_js_error)?)
        };
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_protection(protection);
        Ok(())
    }

    #[wasm_bindgen(js_name = setProtectedRanges)]
    pub fn set_protected_ranges(&self, ranges: JsValue) -> Result<(), JsError> {
        let inputs: Vec<WasmProtectedRangeInput> =
            serde_wasm_bindgen::from_value(ranges).map_err(to_js_error)?;
        let ranges = inputs
            .into_iter()
            .map(WasmProtectedRangeInput::into_core)
            .collect::<Result<Vec<_>, _>>()
            .map_err(to_js_error)?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_protected_ranges(ranges);
        Ok(())
    }

    #[wasm_bindgen(js_name = getCell)]
    pub fn get_cell(&self, address: &str) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        Ok(CellValue {
            inner: ws.get_value_at(addr.row, addr.col),
        })
    }

    #[wasm_bindgen(js_name = getCellAt)]
    pub fn get_cell_at(&self, row: u32, col: u32) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(CellValue {
            inner: ws.get_value_at(row, col as u16),
        })
    }

    #[wasm_bindgen(js_name = getCalculatedValue)]
    pub fn get_calculated_value(&self, address: &str) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        let value = ws
            .get_calculated_value_at(addr.row, addr.col)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);
        Ok(CellValue { inner: value })
    }

    #[wasm_bindgen(js_name = getCalculatedValueAt)]
    pub fn get_calculated_value_at(&self, row: u32, col: u32) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let value = ws
            .get_calculated_value_at(row, col as u16)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);
        Ok(CellValue { inner: value })
    }

    #[wasm_bindgen(js_name = usedRange)]
    pub fn used_range(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.used_range() {
            Some(range) => {
                let arr = js_sys::Array::new();
                arr.push(&JsValue::from(range.start.row));
                arr.push(&JsValue::from(range.start.col));
                arr.push(&JsValue::from(range.end.row));
                arr.push(&JsValue::from(range.end.col));
                Ok(arr.into())
            }
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = setRowHeight)]
    pub fn set_row_height(&self, row: u32, height: f64) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_row_height(row, height);
        Ok(())
    }

    #[wasm_bindgen(js_name = setColumnWidth)]
    pub fn set_column_width(&self, col: u32, width: f64) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_column_width(col as u16, width);
        Ok(())
    }

    #[wasm_bindgen(js_name = getRowHeight)]
    pub fn get_row_height(&self, row: u32) -> Result<Option<f64>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.custom_row_heights().get(&row).copied())
    }

    #[wasm_bindgen(js_name = getColumnWidth)]
    pub fn get_column_width(&self, col: u32) -> Result<Option<f64>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.custom_column_widths().get(&(col as u16)).copied())
    }

    #[wasm_bindgen(js_name = mergeCells)]
    pub fn merge_cells(&self, range_str: &str) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        ws.merge_cells(&range).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = unmergeCells)]
    pub fn unmerge_cells(&self, range_str: &str) -> Result<bool, JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        Ok(ws.unmerge_cells(&range))
    }

    #[wasm_bindgen(js_name = getImageAt)]
    pub fn get_image_at(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        match ws.get_image_at(row, col as u16) {
            Some(info) => to_js_value(&WasmImageInfo::from(info)),
            None => Ok(JsValue::NULL),
        }
    }
}

#[wasm_bindgen]
pub struct Workbook {
    pub(crate) inner: Rc<RefCell<CoreWorkbook>>,
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: Rc::new(RefCell::new(CoreWorkbook::new())),
        }
    }

    #[wasm_bindgen(js_name = loadCsvString)]
    pub fn load_csv_string(csv: &str) -> Result<Workbook, JsError> {
        let reader = Cursor::new(csv.as_bytes());
        let ws =
            duke_sheets_csv::CsvReader::read(reader, &duke_sheets_csv::CsvReadOptions::default())
                .map_err(to_js_error)?;
        let mut wb = CoreWorkbook::empty();
        wb.add_existing_worksheet(ws).map_err(to_js_error)?;
        Ok(Self {
            inner: Rc::new(RefCell::new(wb)),
        })
    }

    #[wasm_bindgen(js_name = fromCsvString)]
    pub fn from_csv_string(csv: &str) -> Result<Workbook, JsError> {
        Self::load_csv_string(csv)
    }

    /// Load a workbook from bytes, auto-detecting the format (XLSX or XLS).
    #[wasm_bindgen(js_name = fromBytes)]
    pub fn from_bytes(data: &[u8]) -> Result<Workbook, JsError> {
        use duke_sheets::WorkbookExt;
        let wb = duke_sheets_core::Workbook::from_bytes(data)
            .map_err(|e| JsError::new(&format!("Failed to read file: {}", e)))?;
        Ok(Self {
            inner: Rc::new(RefCell::new(wb)),
        })
    }

    /// Load a password-protected workbook from bytes, auto-detecting
    /// the format (XLSX or XLS). `skipIntegrityCheck` skips the HMAC
    /// integrity check on Agile-encrypted files (matches Office
    /// behaviour); defaults to false.
    #[wasm_bindgen(js_name = fromBytesWithPassword)]
    pub fn from_bytes_with_password(
        data: &[u8],
        password: &str,
        skip_integrity_check: Option<bool>,
    ) -> Result<Workbook, JsError> {
        use duke_sheets::{WorkbookExt, WorkbookOpenOptions};
        let mut opts = WorkbookOpenOptions::default().password(password);
        if skip_integrity_check.unwrap_or(false) {
            opts = opts.skip_integrity_check();
        }
        let wb = duke_sheets_core::Workbook::from_bytes_with(data, &opts)
            .map_err(|e| JsError::new(&format!("Failed to read file: {}", e)))?;
        Ok(Self {
            inner: Rc::new(RefCell::new(wb)),
        })
    }

    /// Save XLSX bytes with form-control state synchronized into linked cells,
    /// replacing existing values and formulas in the output.
    #[wasm_bindgen(js_name = saveXlsxBytes)]
    pub fn save_xlsx_bytes(&self) -> Result<Vec<u8>, JsError> {
        let wb = self.inner.borrow();
        let snapshot = wb.synchronized_for_save();
        let wb = snapshot.as_ref().unwrap_or_else(|| &*wb);
        let mut buf = Vec::new();
        XlsxWriter::write(wb, Cursor::new(&mut buf)).map_err(to_js_error)?;
        Ok(buf)
    }

    /// Save the workbook as encrypted XLSX bytes after synchronizing
    /// form-control state into linked cells. `profile` selects
    /// the encryption variant; passing `null`/`undefined` uses the
    /// Agile-256 default. Valid values: `"agile"`, `"standard"`.
    /// `keyBits` and `spinCount` override the defaults where the
    /// profile supports them.
    #[wasm_bindgen(js_name = saveXlsxBytesEncrypted)]
    pub fn save_xlsx_bytes_encrypted(
        &self,
        password: &str,
        profile: Option<String>,
        key_bits: Option<u32>,
        spin_count: Option<u32>,
    ) -> Result<Vec<u8>, JsError> {
        let xlsx_profile = parse_xlsx_profile(profile.as_deref(), key_bits, spin_count)
            .map_err(|e| JsError::new(&e))?;
        let wb = self.inner.borrow();
        let snapshot = wb.synchronized_for_save();
        let wb = snapshot.as_ref().unwrap_or_else(|| &*wb);
        XlsxWriter::write_to_bytes_encrypted(wb, password, &xlsx_profile).map_err(to_js_error)
    }

    /// Save the workbook as encrypted XLS bytes after synchronizing
    /// form-control state into linked cells. `profile` selects
    /// the FilePass variant; `null` defaults to RC4 CryptoAPI 128.
    /// Valid values: `"rc4-cryptoapi"`, `"rc4-legacy"`, `"xor"`.
    /// `keyBits` controls RC4 CryptoAPI key size (40 or 128). XOR is
    /// not certified to interoperate with modern Excel.
    #[wasm_bindgen(js_name = saveXlsBytesEncrypted)]
    pub fn save_xls_bytes_encrypted(
        &self,
        password: &str,
        profile: Option<String>,
        key_bits: Option<u32>,
    ) -> Result<Vec<u8>, JsError> {
        let variant =
            parse_xls_variant(profile.as_deref(), key_bits).map_err(|e| JsError::new(&e))?;
        let wb = self.inner.borrow();
        let snapshot = wb.synchronized_for_save();
        let wb = snapshot.as_ref().unwrap_or_else(|| &*wb);
        duke_sheets_xls::XlsWriter::write_to_bytes_encrypted(wb, password, variant)
            .map_err(to_js_error)
    }

    /// Save XLSB bytes with form-control state synchronized into linked cells,
    /// replacing existing values and formulas in the output.
    #[wasm_bindgen(js_name = saveXlsbBytes)]
    pub fn save_xlsb_bytes(&self) -> Result<Vec<u8>, JsError> {
        let wb = self.inner.borrow();
        let snapshot = wb.synchronized_for_save();
        let wb = snapshot.as_ref().unwrap_or_else(|| &*wb);
        let mut buf = Vec::new();
        XlsbWriter::write(wb, Cursor::new(&mut buf)).map_err(to_js_error)?;
        Ok(buf)
    }

    /// Save the first sheet as a CSV string, with form-control state
    /// synchronized into linked cells in the output.
    #[wasm_bindgen(js_name = saveCsvString)]
    pub fn save_csv_string(&self) -> Result<String, JsError> {
        let wb = self.inner.borrow();
        let snapshot = wb.synchronized_for_save();
        let wb = snapshot.as_ref().unwrap_or_else(|| &*wb);
        let ws = wb
            .worksheet(0)
            .ok_or_else(|| JsError::new("No worksheets to save"))?;
        let mut buf = Vec::new();
        duke_sheets_csv::CsvWriter::write(
            ws,
            &mut buf,
            &duke_sheets_csv::CsvWriteOptions::default(),
        )
        .map_err(to_js_error)?;
        String::from_utf8(buf).map_err(to_js_error)
    }

    #[wasm_bindgen(getter, js_name = sheetCount)]
    pub fn sheet_count(&self) -> Result<usize, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.sheet_count())
    }

    #[wasm_bindgen(getter, js_name = sheetNames)]
    pub fn sheet_names(&self) -> Result<Vec<String>, JsError> {
        let wb = self.inner.borrow();
        Ok((0..wb.sheet_count())
            .filter_map(|i| wb.worksheet(i).map(|ws| ws.name().to_string()))
            .collect())
    }

    #[wasm_bindgen(js_name = getSheet)]
    pub fn get_sheet(&self, index: usize) -> Result<Worksheet, JsError> {
        let wb = self.inner.borrow();
        if index >= wb.sheet_count() {
            return Err(JsError::new(&format!("Sheet index {} out of range", index)));
        }
        drop(wb);
        Ok(Worksheet {
            workbook: Rc::clone(&self.inner),
            sheet_index: index,
        })
    }

    #[wasm_bindgen(js_name = getSheetByName)]
    pub fn get_sheet_by_name(&self, name: &str) -> Result<Worksheet, JsError> {
        let wb = self.inner.borrow();
        let index = wb
            .sheet_index(name)
            .ok_or_else(|| JsError::new(&format!("Sheet '{}' not found", name)))?;
        drop(wb);
        Ok(Worksheet {
            workbook: Rc::clone(&self.inner),
            sheet_index: index,
        })
    }

    #[wasm_bindgen(js_name = addSheet)]
    pub fn add_sheet(&self, name: &str) -> Result<usize, JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.add_worksheet_with_name(name).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setWorkbookProtection)]
    pub fn set_workbook_protection(&self, protection: JsValue) -> Result<(), JsError> {
        let protection = if protection.is_null() || protection.is_undefined() {
            None
        } else {
            let input: WasmWorkbookProtectionInput =
                serde_wasm_bindgen::from_value(protection).map_err(to_js_error)?;
            Some(input.into_core().map_err(to_js_error)?)
        };
        let mut wb = self.inner.borrow_mut();
        wb.set_workbook_protection(protection);
        Ok(())
    }

    #[wasm_bindgen(js_name = removeSheet)]
    pub fn remove_sheet(&self, index: usize) -> Result<(), JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.remove_worksheet(index).map(|_| ()).map_err(to_js_error)
    }

    pub fn calculate(&self, options: Option<JsValue>) -> Result<JsValue, JsError> {
        let mut wb = self.inner.borrow_mut();
        let options = if let Some(options) = options {
            // Extract JS callback functions before serde deserialization (which consumes the JsValue).
            // Serde can't deserialize JS functions, so we pull them out via Reflect::get.
            let web_service_js_fn =
                js_sys::Reflect::get(&options, &JsValue::from_str("webServiceFn"))
                    .ok()
                    .and_then(|v| v.dyn_into::<js_sys::Function>().ok());
            let rtd_js_fn = js_sys::Reflect::get(&options, &JsValue::from_str("rtdFn"))
                .ok()
                .and_then(|v| v.dyn_into::<js_sys::Function>().ok());
            let external_js_fn = js_sys::Reflect::get(&options, &JsValue::from_str("externalFn"))
                .ok()
                .and_then(|v| v.dyn_into::<js_sys::Function>().ok())
                .map(|f| (f, false))
                .or_else(|| {
                    js_sys::Reflect::get(&options, &JsValue::from_str("externalFnFn"))
                        .ok()
                        .and_then(|v| v.dyn_into::<js_sys::Function>().ok())
                        .map(|f| (f, true))
                });

            let js_opts: JsCalculationOptions =
                serde_wasm_bindgen::from_value(options).map_err(to_js_error)?;

            // Build web_service_fn from callback
            let web_service_fn = web_service_js_fn.map(|js_fn| {
                let wrapper = SendSyncFunction(js_fn);
                Arc::new(move |url: &str| -> Option<String> {
                    let result = wrapper
                        .call1(&JsValue::NULL, &JsValue::from_str(url))
                        .ok()?;
                    result.as_string()
                }) as Arc<dyn Fn(&str) -> Option<String> + Send + Sync>
            });

            // Build rtd_fn from callback
            let rtd_fn = rtd_js_fn.map(|js_fn| {
                let wrapper = SendSyncFunction(js_fn);
                Arc::new(
                    move |prog_id: &str, server: &str, topics: &[String]| -> Option<String> {
                        let topics_arr = js_sys::Array::new();
                        for t in topics {
                            topics_arr.push(&JsValue::from_str(t));
                        }
                        let result = wrapper
                            .call3(
                                &JsValue::NULL,
                                &JsValue::from_str(prog_id),
                                &JsValue::from_str(server),
                                &topics_arr.into(),
                            )
                            .ok()?;
                        result.as_string()
                    },
                )
                    as Arc<dyn Fn(&str, &str, &[String]) -> Option<String> + Send + Sync>
            });

            // Build external_fn from callback: externalFn(book, name, args[]) -> number|string|bool|null
            let external_fn = external_js_fn.map(|(js_fn, legacy_two_arg)| {
                let wrapper = SendSyncFunction(js_fn);
                Arc::new(
                    move |book: &str, name: &str, args: &[String]| -> Option<FormulaValue> {
                        let args_arr = js_sys::Array::new();
                        for a in args {
                            args_arr.push(&JsValue::from_str(a));
                        }
                        let result = if legacy_two_arg {
                            wrapper.call2(
                                &JsValue::NULL,
                                &JsValue::from_str(name),
                                &args_arr.into(),
                            )
                        } else {
                            wrapper.call3(
                                &JsValue::NULL,
                                &JsValue::from_str(book),
                                &JsValue::from_str(name),
                                &args_arr.into(),
                            )
                        }
                        .ok()?;
                        js_value_to_formula_value(result)
                    },
                )
                    as Arc<dyn Fn(&str, &str, &[String]) -> Option<FormulaValue> + Send + Sync>
            });

            CalculationOptions {
                iterative: js_opts.iterative.unwrap_or(false),
                max_iterations: js_opts.max_iterations.unwrap_or(100),
                max_change: js_opts.max_change.unwrap_or(0.001),
                force_full_calculation: js_opts.force_full_calculation.unwrap_or(true),
                calculate_volatile: js_opts.calculate_volatile.unwrap_or(true),
                sheets: js_opts.sheets.unwrap_or_default(),
                max_threads: js_opts.max_threads,
                web_service_fn,
                rtd_fn,
                external_fn,
            }
        } else {
            CalculationOptions::default()
        };
        let stats = wb.calculate_with_options(&options).map_err(to_js_error)?;
        to_js_value(&WasmCalculationStats::from(&stats))
    }

    #[wasm_bindgen(js_name = defineName)]
    pub fn define_name(&self, name: &str, refers_to: &str) -> Result<(), JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.define_name(name, refers_to).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = getNamedRange)]
    pub fn get_named_range(&self, name: &str) -> Result<Option<String>, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.get_named_range(name, 0).map(|nr| nr.refers_to.clone()))
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

/// Convert a JS value returned by an `externalFnFn` callback into a typed FormulaValue.
/// `null`/`undefined` -> `None` (yields `#N/A`). Booleans take priority over numbers.
fn js_value_to_formula_value(value: JsValue) -> Option<FormulaValue> {
    if value.is_null() || value.is_undefined() {
        None
    } else if let Some(b) = value.as_bool() {
        Some(FormulaValue::Boolean(b))
    } else if let Some(n) = value.as_f64() {
        Some(FormulaValue::Number(n))
    } else if let Some(s) = value.as_string() {
        Some(FormulaValue::String(s))
    } else {
        None
    }
}

fn js_to_cell_value(value: JsValue) -> Result<CoreCellValue, JsError> {
    if value.is_null() || value.is_undefined() {
        Ok(CoreCellValue::Empty)
    } else if let Some(b) = value.as_bool() {
        Ok(CoreCellValue::Boolean(b))
    } else if let Some(n) = value.as_f64() {
        Ok(CoreCellValue::Number(n))
    } else if let Some(s) = value.as_string() {
        Ok(CoreCellValue::string(s))
    } else {
        Err(JsError::new(
            "Cell value must be null, boolean, number, or string",
        ))
    }
}

#[wasm_bindgen(start)]
pub fn init() {}
