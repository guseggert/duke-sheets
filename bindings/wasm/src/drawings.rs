use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;

use duke_sheets_chart::{
    CellMarker, ChildTransform, DrawingAnchor, EditAs, EmbeddedImage, GroupTransform, ImageFormat,
};
use duke_sheets_core as core;

use crate::{to_js_error, to_js_value, Workbook, Worksheet};
use crate::types::{WasmChart, WasmChartEx};

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmDrawingMeta {
    #[serde(skip_serializing_if = "Option::is_none")]
    name: Option<String>,
    hidden: bool,
    locked: bool,
    printable: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    alt_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    title: Option<String>,
}

impl From<&core::DrawingMeta> for WasmDrawingMeta {
    fn from(meta: &core::DrawingMeta) -> Self {
        Self {
            name: meta.name.clone(),
            hidden: meta.hidden,
            locked: meta.locked,
            printable: meta.printable,
            alt_text: meta.alt_text.clone(),
            title: meta.title.clone(),
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmDrawingMetaInput {
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    hidden: Option<bool>,
    #[serde(default)]
    locked: Option<bool>,
    #[serde(default)]
    printable: Option<bool>,
    #[serde(default)]
    alt_text: Option<String>,
    #[serde(default)]
    title: Option<String>,
}

impl WasmDrawingMetaInput {
    fn into_core(self, comment: bool) -> core::DrawingMeta {
        core::DrawingMeta {
            name: self.name,
            hidden: self.hidden.unwrap_or(comment),
            locked: self.locked.unwrap_or(true),
            printable: self.printable.unwrap_or(true),
            alt_text: self.alt_text,
            title: self.title,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmCellMarker {
    col: u16,
    row: u32,
    #[serde(default)]
    col_offset_emu: i64,
    #[serde(default)]
    row_offset_emu: i64,
}

impl From<&CellMarker> for WasmCellMarker {
    fn from(marker: &CellMarker) -> Self {
        Self {
            col: marker.col,
            row: marker.row,
            col_offset_emu: marker.col_offset_emu,
            row_offset_emu: marker.row_offset_emu,
        }
    }
}

impl From<WasmCellMarker> for CellMarker {
    fn from(marker: WasmCellMarker) -> Self {
        Self {
            col: marker.col,
            row: marker.row,
            col_offset_emu: marker.col_offset_emu,
            row_offset_emu: marker.row_offset_emu,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum WasmDrawingAnchor {
    TwoCell {
        from: WasmCellMarker,
        to: WasmCellMarker,
        #[serde(default)]
        edit_as: Option<String>,
    },
    OneCell {
        from: WasmCellMarker,
        width_emu: i64,
        height_emu: i64,
    },
    Absolute {
        x_emu: i64,
        y_emu: i64,
        width_emu: i64,
        height_emu: i64,
    },
}

impl From<&DrawingAnchor> for WasmDrawingAnchor {
    fn from(anchor: &DrawingAnchor) -> Self {
        match anchor {
            DrawingAnchor::TwoCell { from, to, edit_as } => Self::TwoCell {
                from: WasmCellMarker::from(from),
                to: WasmCellMarker::from(to),
                edit_as: edit_as.as_ref().map(|edit_as| match edit_as {
                    EditAs::TwoCell => "twoCell".to_string(),
                    EditAs::OneCell => "oneCell".to_string(),
                    EditAs::Absolute => "absolute".to_string(),
                }),
            },
            DrawingAnchor::OneCell {
                from,
                width_emu,
                height_emu,
            } => Self::OneCell {
                from: WasmCellMarker::from(from),
                width_emu: *width_emu,
                height_emu: *height_emu,
            },
            DrawingAnchor::Absolute {
                x_emu,
                y_emu,
                width_emu,
                height_emu,
            } => Self::Absolute {
                x_emu: *x_emu,
                y_emu: *y_emu,
                width_emu: *width_emu,
                height_emu: *height_emu,
            },
        }
    }
}

impl TryFrom<WasmDrawingAnchor> for DrawingAnchor {
    type Error = String;

    fn try_from(anchor: WasmDrawingAnchor) -> Result<Self, Self::Error> {
        Ok(match anchor {
            WasmDrawingAnchor::TwoCell { from, to, edit_as } => DrawingAnchor::TwoCell {
                from: from.into(),
                to: to.into(),
                edit_as: edit_as
                    .map(|edit_as| match edit_as.as_str() {
                        "twoCell" => Ok(EditAs::TwoCell),
                        "oneCell" => Ok(EditAs::OneCell),
                        "absolute" => Ok(EditAs::Absolute),
                        other => Err(format!("invalid drawing editAs {other:?}")),
                    })
                    .transpose()?,
            },
            WasmDrawingAnchor::OneCell {
                from,
                width_emu,
                height_emu,
            } => DrawingAnchor::OneCell {
                from: from.into(),
                width_emu,
                height_emu,
            },
            WasmDrawingAnchor::Absolute {
                x_emu,
                y_emu,
                width_emu,
                height_emu,
            } => DrawingAnchor::Absolute {
                x_emu,
                y_emu,
                width_emu,
                height_emu,
            },
        })
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmChildTransform {
    #[serde(default)]
    x_emu: i64,
    #[serde(default)]
    y_emu: i64,
    #[serde(default)]
    cx_emu: i64,
    #[serde(default)]
    cy_emu: i64,
    #[serde(default)]
    rotation: i32,
    #[serde(default)]
    flip_h: bool,
    #[serde(default)]
    flip_v: bool,
}

impl From<&ChildTransform> for WasmChildTransform {
    fn from(transform: &ChildTransform) -> Self {
        Self {
            x_emu: transform.x_emu,
            y_emu: transform.y_emu,
            cx_emu: transform.cx_emu,
            cy_emu: transform.cy_emu,
            rotation: transform.rotation,
            flip_h: transform.flip_h,
            flip_v: transform.flip_v,
        }
    }
}

impl From<WasmChildTransform> for ChildTransform {
    fn from(transform: WasmChildTransform) -> Self {
        Self {
            x_emu: transform.x_emu,
            y_emu: transform.y_emu,
            cx_emu: transform.cx_emu,
            cy_emu: transform.cy_emu,
            rotation: transform.rotation,
            flip_h: transform.flip_h,
            flip_v: transform.flip_v,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct WasmGroupTransform {
    #[serde(default)]
    x_emu: i64,
    #[serde(default)]
    y_emu: i64,
    #[serde(default)]
    cx_emu: i64,
    #[serde(default)]
    cy_emu: i64,
    #[serde(default)]
    child_x_emu: i64,
    #[serde(default)]
    child_y_emu: i64,
    #[serde(default)]
    child_cx_emu: i64,
    #[serde(default)]
    child_cy_emu: i64,
    #[serde(default)]
    rotation: i32,
    #[serde(default)]
    flip_h: bool,
    #[serde(default)]
    flip_v: bool,
}

impl From<&GroupTransform> for WasmGroupTransform {
    fn from(transform: &GroupTransform) -> Self {
        Self {
            x_emu: transform.x_emu,
            y_emu: transform.y_emu,
            cx_emu: transform.cx_emu,
            cy_emu: transform.cy_emu,
            child_x_emu: transform.child_x_emu,
            child_y_emu: transform.child_y_emu,
            child_cx_emu: transform.child_cx_emu,
            child_cy_emu: transform.child_cy_emu,
            rotation: transform.rotation,
            flip_h: transform.flip_h,
            flip_v: transform.flip_v,
        }
    }
}

impl From<WasmGroupTransform> for GroupTransform {
    fn from(transform: WasmGroupTransform) -> Self {
        Self {
            x_emu: transform.x_emu,
            y_emu: transform.y_emu,
            cx_emu: transform.cx_emu,
            cy_emu: transform.cy_emu,
            child_x_emu: transform.child_x_emu,
            child_y_emu: transform.child_y_emu,
            child_cx_emu: transform.child_cx_emu,
            child_cy_emu: transform.child_cy_emu,
            rotation: transform.rotation,
            flip_h: transform.flip_h,
            flip_v: transform.flip_v,
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(untagged)]
enum WasmDrawingPlacement {
    TopLevel { anchor: WasmDrawingAnchor },
    Child { transform: WasmChildTransform },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "colorType", rename_all = "camelCase", rename_all_fields = "camelCase")]
pub(crate) enum WasmDrawingColor {
    Auto,
    Rgb { r: u8, g: u8, b: u8 },
    Argb { a: u8, r: u8, g: u8, b: u8 },
    Theme { index: u8, tint: f64 },
    Indexed { index: u8 },
}

impl From<&core::Color> for WasmDrawingColor {
    fn from(color: &core::Color) -> Self {
        match color {
            core::Color::Auto => Self::Auto,
            core::Color::Rgb { r, g, b } => Self::Rgb {
                r: *r,
                g: *g,
                b: *b,
            },
            core::Color::Argb { a, r, g, b } => Self::Argb {
                a: *a,
                r: *r,
                g: *g,
                b: *b,
            },
            core::Color::Theme { index, tint } => Self::Theme {
                index: *index,
                tint: *tint,
            },
            core::Color::Indexed(index) => Self::Indexed { index: *index },
        }
    }
}

impl From<WasmDrawingColor> for core::Color {
    fn from(color: WasmDrawingColor) -> Self {
        match color {
            WasmDrawingColor::Auto => Self::Auto,
            WasmDrawingColor::Rgb { r, g, b } => Self::Rgb { r, g, b },
            WasmDrawingColor::Argb { a, r, g, b } => Self::Argb { a, r, g, b },
            WasmDrawingColor::Theme { index, tint } => Self::Theme { index, tint },
            WasmDrawingColor::Indexed { index } => Self::Indexed(index),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct WasmDrawingRunFont {
    #[serde(default)]
    bold: Option<bool>,
    #[serde(default)]
    italic: Option<bool>,
    #[serde(default)]
    size: Option<f64>,
    #[serde(default)]
    color: Option<WasmDrawingColor>,
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    underline: Option<String>,
    #[serde(default)]
    strikethrough: Option<bool>,
    #[serde(default)]
    vertical_align: Option<String>,
    #[serde(default)]
    family: Option<u8>,
    #[serde(default)]
    charset: Option<u8>,
    #[serde(default)]
    scheme: Option<String>,
}

impl From<&core::RunFont> for WasmDrawingRunFont {
    fn from(font: &core::RunFont) -> Self {
        Self {
            bold: font.bold,
            italic: font.italic,
            size: font.size,
            color: font.color.as_ref().map(WasmDrawingColor::from),
            name: font.name.clone(),
            underline: font.underline.map(|underline| match underline {
                core::style::Underline::None => "none".to_string(),
                core::style::Underline::Single => "single".to_string(),
                core::style::Underline::Double => "double".to_string(),
                core::style::Underline::SingleAccounting => "singleAccounting".to_string(),
                core::style::Underline::DoubleAccounting => "doubleAccounting".to_string(),
            }),
            strikethrough: font.strikethrough,
            vertical_align: font.vertical_align.map(|vertical_align| match vertical_align {
                core::style::FontVerticalAlign::Baseline => "baseline".to_string(),
                core::style::FontVerticalAlign::Superscript => "superscript".to_string(),
                core::style::FontVerticalAlign::Subscript => "subscript".to_string(),
            }),
            family: font.family,
            charset: font.charset,
            scheme: font.scheme.clone(),
        }
    }
}

impl TryFrom<WasmDrawingRunFont> for core::RunFont {
    type Error = String;

    fn try_from(font: WasmDrawingRunFont) -> Result<Self, Self::Error> {
        Ok(Self {
            bold: font.bold,
            italic: font.italic,
            size: font.size,
            color: font.color.map(Into::into),
            name: font.name,
            underline: font
                .underline
                .map(|underline| match underline.as_str() {
                    "none" => Ok(core::style::Underline::None),
                    "single" => Ok(core::style::Underline::Single),
                    "double" => Ok(core::style::Underline::Double),
                    "singleAccounting" => Ok(core::style::Underline::SingleAccounting),
                    "doubleAccounting" => Ok(core::style::Underline::DoubleAccounting),
                    other => Err(format!("invalid drawing text underline {other:?}")),
                })
                .transpose()?,
            strikethrough: font.strikethrough,
            vertical_align: font
                .vertical_align
                .map(|vertical_align| match vertical_align.as_str() {
                    "baseline" => Ok(core::style::FontVerticalAlign::Baseline),
                    "superscript" => Ok(core::style::FontVerticalAlign::Superscript),
                    "subscript" => Ok(core::style::FontVerticalAlign::Subscript),
                    other => Err(format!("invalid drawing text verticalAlign {other:?}")),
                })
                .transpose()?,
            family: font.family,
            charset: font.charset,
            scheme: font.scheme,
        })
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmDrawingTextRun {
    text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    font: Option<WasmDrawingRunFont>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct WasmDrawingText {
    #[serde(default)]
    runs: Vec<WasmDrawingTextRun>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    horizontal_alignment: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    vertical_alignment: Option<String>,
}

impl From<&core::DrawingText> for WasmDrawingText {
    fn from(text: &core::DrawingText) -> Self {
        Self {
            runs: text
                .runs
                .iter()
                .map(|run| WasmDrawingTextRun {
                    text: run.text.clone(),
                    font: run.font.as_ref().map(WasmDrawingRunFont::from),
                })
                .collect(),
            horizontal_alignment: text.horizontal_alignment.map(|alignment| match alignment {
                core::HorizontalAlignment::General => "general".to_string(),
                core::HorizontalAlignment::Left => "left".to_string(),
                core::HorizontalAlignment::Center => "center".to_string(),
                core::HorizontalAlignment::Right => "right".to_string(),
                core::HorizontalAlignment::Fill => "fill".to_string(),
                core::HorizontalAlignment::Justify => "justify".to_string(),
                core::HorizontalAlignment::CenterContinuous => "centerContinuous".to_string(),
                core::HorizontalAlignment::Distributed => "distributed".to_string(),
            }),
            vertical_alignment: text.vertical_alignment.map(|alignment| match alignment {
                core::VerticalAlignment::Top => "top".to_string(),
                core::VerticalAlignment::Center => "center".to_string(),
                core::VerticalAlignment::Bottom => "bottom".to_string(),
                core::VerticalAlignment::Justify => "justify".to_string(),
                core::VerticalAlignment::Distributed => "distributed".to_string(),
            }),
        }
    }
}

impl TryFrom<WasmDrawingText> for core::DrawingText {
    type Error = String;

    fn try_from(text: WasmDrawingText) -> Result<Self, Self::Error> {
        let horizontal_alignment = text
            .horizontal_alignment
            .map(|alignment| match alignment.as_str() {
                "general" => Ok(core::HorizontalAlignment::General),
                "left" => Ok(core::HorizontalAlignment::Left),
                "center" => Ok(core::HorizontalAlignment::Center),
                "right" => Ok(core::HorizontalAlignment::Right),
                "fill" => Ok(core::HorizontalAlignment::Fill),
                "justify" => Ok(core::HorizontalAlignment::Justify),
                "centerContinuous" => Ok(core::HorizontalAlignment::CenterContinuous),
                "distributed" => Ok(core::HorizontalAlignment::Distributed),
                other => Err(format!("invalid drawing text horizontalAlignment {other:?}")),
            })
            .transpose()?;
        let vertical_alignment = text
            .vertical_alignment
            .map(|alignment| match alignment.as_str() {
                "top" => Ok(core::VerticalAlignment::Top),
                "center" => Ok(core::VerticalAlignment::Center),
                "bottom" => Ok(core::VerticalAlignment::Bottom),
                "justify" => Ok(core::VerticalAlignment::Justify),
                "distributed" => Ok(core::VerticalAlignment::Distributed),
                other => Err(format!("invalid drawing text verticalAlignment {other:?}")),
            })
            .transpose()?;
        let runs = text
            .runs
            .into_iter()
            .map(|run| {
                Ok(core::RichTextRun {
                    text: run.text,
                    font: run.font.map(TryInto::try_into).transpose()?,
                })
            })
            .collect::<Result<Vec<_>, String>>()?;
        Ok(Self {
            runs,
            horizontal_alignment,
            vertical_alignment,
        })
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
enum WasmCheckState {
    Unchecked,
    Checked,
    Mixed,
}

impl From<core::CheckState> for WasmCheckState {
    fn from(state: core::CheckState) -> Self {
        match state {
            core::CheckState::Unchecked => Self::Unchecked,
            core::CheckState::Checked => Self::Checked,
            core::CheckState::Mixed => Self::Mixed,
        }
    }
}

impl From<WasmCheckState> for core::CheckState {
    fn from(state: WasmCheckState) -> Self {
        match state {
            WasmCheckState::Unchecked => Self::Unchecked,
            WasmCheckState::Checked => Self::Checked,
            WasmCheckState::Mixed => Self::Mixed,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
enum WasmListSelection {
    Single,
    Multi,
    Extend,
}

impl From<core::ListSelection> for WasmListSelection {
    fn from(selection: core::ListSelection) -> Self {
        match selection {
            core::ListSelection::Single => Self::Single,
            core::ListSelection::Multi => Self::Multi,
            core::ListSelection::Extend => Self::Extend,
        }
    }
}

impl From<WasmListSelection> for core::ListSelection {
    fn from(selection: WasmListSelection) -> Self {
        match selection {
            WasmListSelection::Single => Self::Single,
            WasmListSelection::Multi => Self::Multi,
            WasmListSelection::Extend => Self::Extend,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum WasmFormControlKind {
    Button {
        caption: WasmDrawingText,
    },
    Checkbox {
        caption: WasmDrawingText,
        state: WasmCheckState,
        #[serde(default)]
        cell_link: Option<String>,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    OptionButton {
        caption: WasmDrawingText,
        state: WasmCheckState,
        #[serde(default)]
        cell_link: Option<String>,
        #[serde(default)]
        first_in_group: bool,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    Label {
        caption: WasmDrawingText,
    },
    GroupBox {
        caption: WasmDrawingText,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    ListBox {
        #[serde(default)]
        input_range: Option<String>,
        #[serde(default)]
        cell_link: Option<String>,
        selection: WasmListSelection,
        #[serde(default)]
        selected: Vec<u16>,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    Dropdown {
        #[serde(default)]
        input_range: Option<String>,
        #[serde(default)]
        cell_link: Option<String>,
        #[serde(default)]
        selected: Option<u16>,
        lines: u16,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    Scrollbar {
        value: u16,
        min: u16,
        max: u16,
        increment: u16,
        page: u16,
        #[serde(default)]
        horizontal: bool,
        #[serde(default)]
        cell_link: Option<String>,
    },
    Spinner {
        value: u16,
        min: u16,
        max: u16,
        increment: u16,
        #[serde(default)]
        cell_link: Option<String>,
    },
    Unknown {
        object_type: String,
        #[serde(default)]
        legacy_object_type: Option<u16>,
        #[serde(default)]
        caption: WasmDrawingText,
        /// Internal passthrough of unmodeled XLSX `formControlPr`
        /// attributes; echoed back unchanged on write.
        #[serde(default)]
        raw_properties: Vec<(String, String)>,
        /// Internal passthrough of the original BIFF OBJ body (byte
        /// array in JS), required for XLS rewrite.
        #[serde(default, deserialize_with = "de_opt_bytes")]
        raw_obj: Option<Vec<u8>>,
    },
}

impl From<&core::FormControlKind> for WasmFormControlKind {
    fn from(kind: &core::FormControlKind) -> Self {
        match kind {
            core::FormControlKind::Button { caption } => Self::Button {
                caption: WasmDrawingText::from(caption),
            },
            core::FormControlKind::Checkbox {
                caption,
                state,
                cell_link,
                no_3d,
            } => Self::Checkbox {
                caption: WasmDrawingText::from(caption),
                state: (*state).into(),
                cell_link: cell_link.clone(),
                no_3d: *no_3d,
            },
            core::FormControlKind::OptionButton {
                caption,
                state,
                cell_link,
                first_in_group,
                no_3d,
            } => Self::OptionButton {
                caption: WasmDrawingText::from(caption),
                state: (*state).into(),
                cell_link: cell_link.clone(),
                first_in_group: *first_in_group,
                no_3d: *no_3d,
            },
            core::FormControlKind::Label { caption } => Self::Label {
                caption: WasmDrawingText::from(caption),
            },
            core::FormControlKind::GroupBox { caption, no_3d } => Self::GroupBox {
                caption: WasmDrawingText::from(caption),
                no_3d: *no_3d,
            },
            core::FormControlKind::ListBox {
                input_range,
                cell_link,
                selection,
                selected,
                no_3d,
            } => Self::ListBox {
                input_range: input_range.clone(),
                cell_link: cell_link.clone(),
                selection: (*selection).into(),
                selected: selected.clone(),
                no_3d: *no_3d,
            },
            core::FormControlKind::Dropdown {
                input_range,
                cell_link,
                selected,
                lines,
                no_3d,
            } => Self::Dropdown {
                input_range: input_range.clone(),
                cell_link: cell_link.clone(),
                selected: *selected,
                lines: *lines,
                no_3d: *no_3d,
            },
            core::FormControlKind::Scrollbar {
                value,
                min,
                max,
                increment,
                page,
                horizontal,
                cell_link,
            } => Self::Scrollbar {
                value: *value,
                min: *min,
                max: *max,
                increment: *increment,
                page: *page,
                horizontal: *horizontal,
                cell_link: cell_link.clone(),
            },
            core::FormControlKind::Spinner {
                value,
                min,
                max,
                increment,
                cell_link,
            } => Self::Spinner {
                value: *value,
                min: *min,
                max: *max,
                increment: *increment,
                cell_link: cell_link.clone(),
            },
            core::FormControlKind::Unknown {
                object_type,
                legacy_object_type,
                caption,
                raw_properties,
                raw_obj,
            } => Self::Unknown {
                object_type: object_type.clone(),
                legacy_object_type: *legacy_object_type,
                caption: WasmDrawingText::from(caption),
                raw_properties: raw_properties.clone(),
                raw_obj: raw_obj.clone(),
            },
        }
    }
}

impl TryFrom<WasmFormControlKind> for core::FormControlKind {
    type Error = String;

    fn try_from(kind: WasmFormControlKind) -> Result<Self, Self::Error> {
        Ok(match kind {
            WasmFormControlKind::Button { caption } => Self::Button {
                caption: caption.try_into()?,
            },
            WasmFormControlKind::Checkbox {
                caption,
                state,
                cell_link,
                no_3d,
            } => Self::Checkbox {
                caption: caption.try_into()?,
                state: state.into(),
                cell_link,
                no_3d,
            },
            WasmFormControlKind::OptionButton {
                caption,
                state,
                cell_link,
                no_3d,
                ..
            } => Self::OptionButton {
                caption: caption.try_into()?,
                state: state.into(),
                cell_link,
                first_in_group: false,
                no_3d,
            },
            WasmFormControlKind::Label { caption } => Self::Label {
                caption: caption.try_into()?,
            },
            WasmFormControlKind::GroupBox { caption, no_3d } => Self::GroupBox {
                caption: caption.try_into()?,
                no_3d,
            },
            WasmFormControlKind::ListBox {
                input_range,
                cell_link,
                selection,
                selected,
                no_3d,
            } => Self::ListBox {
                input_range,
                cell_link,
                selection: selection.into(),
                selected,
                no_3d,
            },
            WasmFormControlKind::Dropdown {
                input_range,
                cell_link,
                selected,
                lines,
                no_3d,
            } => Self::Dropdown {
                input_range,
                cell_link,
                selected,
                lines,
                no_3d,
            },
            WasmFormControlKind::Scrollbar {
                value,
                min,
                max,
                increment,
                page,
                horizontal,
                cell_link,
            } => Self::Scrollbar {
                value,
                min,
                max,
                increment,
                page,
                horizontal,
                cell_link,
            },
            WasmFormControlKind::Spinner {
                value,
                min,
                max,
                increment,
                cell_link,
            } => Self::Spinner {
                value,
                min,
                max,
                increment,
                cell_link,
            },
            WasmFormControlKind::Unknown {
                object_type,
                legacy_object_type,
                caption,
                raw_properties,
                raw_obj,
            } => Self::Unknown {
                object_type,
                legacy_object_type,
                caption: caption.try_into()?,
                raw_properties,
                raw_obj,
            },
        })
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmFormControl {
    kind: WasmFormControlKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    macro_name: Option<String>,
    /// Unmodeled VML `ClientData` children (byte arrays in JS)
    /// preserved on any control kind; echoed back unchanged on write.
    #[serde(default, deserialize_with = "de_bytes_vec")]
    raw_client_data: Vec<Vec<u8>>,
}

impl From<&core::FormControl> for WasmFormControl {
    fn from(control: &core::FormControl) -> Self {
        Self {
            kind: WasmFormControlKind::from(&control.kind),
            macro_name: control.macro_name.clone(),
            raw_client_data: control.raw_client_data.clone(),
        }
    }
}

impl TryFrom<WasmFormControl> for core::FormControl {
    type Error = String;

    fn try_from(control: WasmFormControl) -> Result<Self, Self::Error> {
        let result = Self {
            kind: control.kind.try_into()?,
            macro_name: control.macro_name,
            raw_client_data: control.raw_client_data,
        };
        result.validate().map_err(|error| error.to_string())?;
        Ok(result)
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmImageMetadata {
    format: String,
    #[serde(default)]
    media_path: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    svg_media_path: Option<String>,
    width_emu: i64,
    height_emu: i64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    rotation: Option<i32>,
    #[serde(default)]
    flip_h: bool,
    #[serde(default)]
    flip_v: bool,
}

impl From<&EmbeddedImage> for WasmImageMetadata {
    fn from(image: &EmbeddedImage) -> Self {
        Self {
            format: image.format.as_str().to_string(),
            media_path: image.media_path.clone(),
            svg_media_path: image.svg_media_path.clone(),
            width_emu: image.width_emu,
            height_emu: image.height_emu,
            rotation: image.rotation,
            flip_h: image.flip_h,
            flip_v: image.flip_v,
        }
    }
}

/// Deserialize a JS byte payload as either bytes or a sequence.
///
/// Drawing inputs pass through serde's internal buffering (tagged
/// enums and flatten), which captures a `Uint8Array` as bytes; plain
/// `Vec<u8>` only accepts sequences and would reject it.
fn de_bytes<'de, D: serde::Deserializer<'de>>(deserializer: D) -> Result<Vec<u8>, D::Error> {
    struct BytesVisitor;
    impl<'de> serde::de::Visitor<'de> for BytesVisitor {
        type Value = Vec<u8>;

        fn expecting(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
            f.write_str("bytes or a sequence of byte values")
        }

        fn visit_bytes<E: serde::de::Error>(self, v: &[u8]) -> Result<Self::Value, E> {
            Ok(v.to_vec())
        }

        fn visit_byte_buf<E: serde::de::Error>(self, v: Vec<u8>) -> Result<Self::Value, E> {
            Ok(v)
        }

        fn visit_seq<A: serde::de::SeqAccess<'de>>(
            self,
            mut seq: A,
        ) -> Result<Self::Value, A::Error> {
            let mut out = Vec::with_capacity(seq.size_hint().unwrap_or(0));
            while let Some(byte) = seq.next_element::<u8>()? {
                out.push(byte);
            }
            Ok(out)
        }
    }
    deserializer.deserialize_any(BytesVisitor)
}

/// A byte payload wrapper deserializing through [`de_bytes`].
#[derive(Debug, Clone)]
struct JsBytes(Vec<u8>);

impl<'de> Deserialize<'de> for JsBytes {
    fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        de_bytes(deserializer).map(JsBytes)
    }
}

fn de_opt_bytes<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> Result<Option<Vec<u8>>, D::Error> {
    let value = Option::<JsBytes>::deserialize(deserializer)?;
    Ok(value.map(|bytes| bytes.0))
}

fn de_bytes_vec<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> Result<Vec<Vec<u8>>, D::Error> {
    let value = Vec::<JsBytes>::deserialize(deserializer)?;
    Ok(value.into_iter().map(|bytes| bytes.0).collect())
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmImageInput {
    format: String,
    #[serde(default)]
    media_path: String,
    #[serde(default)]
    svg_media_path: Option<String>,
    width_emu: i64,
    height_emu: i64,
    #[serde(default)]
    rotation: Option<i32>,
    #[serde(default)]
    flip_h: bool,
    #[serde(default)]
    flip_v: bool,
    #[serde(deserialize_with = "de_bytes")]
    data: Vec<u8>,
    #[serde(default, deserialize_with = "de_opt_bytes")]
    svg_data: Option<Vec<u8>>,
}

impl TryFrom<WasmImageInput> for EmbeddedImage {
    type Error = String;

    fn try_from(image: WasmImageInput) -> Result<Self, Self::Error> {
        let format = ImageFormat::from_extension(&image.format)
            .ok_or_else(|| format!("unsupported image format {:?}", image.format))?;
        Ok(Self {
            format,
            media_path: image.media_path,
            svg_media_path: image.svg_media_path,
            width_emu: image.width_emu,
            height_emu: image.height_emu,
            rotation: image.rotation,
            flip_h: image.flip_h,
            flip_v: image.flip_v,
            data: image.data,
            svg_data: image.svg_data,
        })
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum WasmShapeFill {
    None,
    Solid { color: WasmDrawingColor },
}

impl From<&core::ShapeFill> for WasmShapeFill {
    fn from(fill: &core::ShapeFill) -> Self {
        match fill {
            core::ShapeFill::None => Self::None,
            core::ShapeFill::Solid(color) => Self::Solid {
                color: WasmDrawingColor::from(color),
            },
        }
    }
}

impl From<WasmShapeFill> for core::ShapeFill {
    fn from(fill: WasmShapeFill) -> Self {
        match fill {
            WasmShapeFill::None => Self::None,
            WasmShapeFill::Solid { color } => Self::Solid(color.into()),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct WasmShapeLine {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    color: Option<WasmDrawingColor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    width_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    dash_style: Option<String>,
    #[serde(default)]
    no_fill: bool,
}

impl From<&core::ShapeLine> for WasmShapeLine {
    fn from(line: &core::ShapeLine) -> Self {
        Self {
            color: line.color.as_ref().map(WasmDrawingColor::from),
            width_emu: line.width_emu,
            dash_style: line.dash_style.clone(),
            no_fill: line.no_fill,
        }
    }
}

impl From<WasmShapeLine> for core::ShapeLine {
    fn from(line: WasmShapeLine) -> Self {
        Self {
            color: line.color.map(Into::into),
            width_emu: line.width_emu,
            dash_style: line.dash_style,
            no_fill: line.no_fill,
        }
    }
}

fn default_shape_geometry() -> String {
    "rect".to_string()
}

fn default_shape_fill() -> WasmShapeFill {
    WasmShapeFill::None
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmShape {
    #[serde(default = "default_shape_geometry")]
    geometry: String,
    #[serde(default = "default_shape_fill")]
    fill: WasmShapeFill,
    #[serde(default)]
    line: WasmShapeLine,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    text: Option<WasmDrawingText>,
    #[serde(default)]
    rotation: i32,
    #[serde(default)]
    flip_h: bool,
    #[serde(default)]
    flip_v: bool,
}

impl From<&core::Shape> for WasmShape {
    fn from(shape: &core::Shape) -> Self {
        let core::ShapeGeometry::Preset(geometry) = &shape.geometry;
        Self {
            geometry: geometry.clone(),
            fill: WasmShapeFill::from(&shape.fill),
            line: WasmShapeLine::from(&shape.line),
            text: shape.text.as_ref().map(WasmDrawingText::from),
            rotation: shape.rotation,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
        }
    }
}

impl TryFrom<WasmShape> for core::Shape {
    type Error = String;

    fn try_from(shape: WasmShape) -> Result<Self, Self::Error> {
        Ok(Self {
            geometry: core::ShapeGeometry::Preset(shape.geometry),
            fill: shape.fill.into(),
            line: shape.line.into(),
            text: shape.text.map(TryInto::try_into).transpose()?,
            rotation: shape.rotation,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
            ..core::Shape::default()
        })
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmDrawingComment {
    row: u32,
    col: u16,
    author: String,
    text: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmRawRelationshipMetadata {
    id: String,
    rel_type: String,
    target: String,
    external: bool,
    has_part: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmRawDrawingMetadata {
    byte_length: usize,
    relationships: Vec<WasmRawRelationshipMetadata>,
}

#[derive(Clone, Deserialize)]
#[serde(tag = "refType", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum WasmChartDataReferenceInput {
    Formula { formula: String },
    Numbers { numbers: Vec<f64> },
    Strings { strings: Vec<String> },
}

impl From<WasmChartDataReferenceInput> for duke_sheets_chart::DataReference {
    fn from(reference: WasmChartDataReferenceInput) -> Self {
        match reference {
            WasmChartDataReferenceInput::Formula { formula } => Self::Formula(formula),
            WasmChartDataReferenceInput::Numbers { numbers } => Self::Numbers(numbers),
            WasmChartDataReferenceInput::Strings { strings } => Self::Strings(strings),
        }
    }
}

#[derive(Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmChartSeriesInput {
    #[serde(default)]
    name: Option<String>,
    values: WasmChartDataReferenceInput,
    #[serde(default)]
    categories: Option<WasmChartDataReferenceInput>,
}

#[derive(Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmChartInput {
    chart_type: String,
    #[serde(default)]
    title: Option<String>,
    #[serde(default)]
    series: Vec<WasmChartSeriesInput>,
    #[serde(default)]
    is_3d: bool,
    #[serde(default)]
    vary_colors: Option<bool>,
    #[serde(default)]
    gap_width: Option<u32>,
    #[serde(default)]
    overlap: Option<i32>,
}

fn chart_type_from_string(value: &str) -> Result<duke_sheets_chart::ChartType, String> {
    use duke_sheets_chart::ChartType;
    Ok(match value {
        "ColumnClustered" => ChartType::ColumnClustered,
        "ColumnStacked" => ChartType::ColumnStacked,
        "ColumnPercentStacked" => ChartType::ColumnPercentStacked,
        "BarClustered" => ChartType::BarClustered,
        "BarStacked" => ChartType::BarStacked,
        "BarPercentStacked" => ChartType::BarPercentStacked,
        "Line" => ChartType::Line,
        "LineStacked" => ChartType::LineStacked,
        "Pie" => ChartType::Pie,
        "PieExploded" => ChartType::PieExploded,
        "Doughnut" => ChartType::Doughnut,
        "Area" => ChartType::Area,
        "AreaStacked" => ChartType::AreaStacked,
        "AreaPercentStacked" => ChartType::AreaPercentStacked,
        "ScatterMarkers" => ChartType::ScatterMarkers,
        "ScatterSmooth" => ChartType::ScatterSmooth,
        "ScatterLines" => ChartType::ScatterLines,
        "Bubble" => ChartType::Bubble,
        "Radar" => ChartType::Radar,
        "Stock" => ChartType::Stock,
        "Surface" => ChartType::Surface,
        other if other.starts_with("Unsupported(") && other.ends_with(')') => {
            ChartType::Unsupported(other[12..other.len() - 1].to_string())
        }
        other => return Err(format!("unsupported chart type {other:?}")),
    })
}

impl TryFrom<WasmChartInput> for duke_sheets_chart::Chart {
    type Error = String;

    fn try_from(input: WasmChartInput) -> Result<Self, Self::Error> {
        let mut chart = Self::new(chart_type_from_string(&input.chart_type)?);
        chart.title = input.title;
        chart.series = input
            .series
            .into_iter()
            .map(|series| {
                let mut core_series =
                    duke_sheets_chart::DataSeries::new(series.values.into());
                core_series.name = series.name;
                core_series.categories = series.categories.map(Into::into);
                core_series
            })
            .collect();
        chart.is_3d = input.is_3d;
        chart.vary_colors = input.vary_colors;
        chart.gap_width = input.gap_width;
        chart.overlap = input.overlap;
        Ok(chart)
    }
}

#[derive(Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmChartExTitleInput {
    #[serde(default)]
    text: Option<String>,
    #[serde(default)]
    position: Option<String>,
    #[serde(default)]
    align: Option<String>,
    #[serde(default)]
    overlay: Option<bool>,
}

#[derive(Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmChartExInput {
    layout: String,
    #[serde(default)]
    version: Option<String>,
    #[serde(default)]
    feature_list: Option<String>,
    #[serde(default)]
    fallback_img: Option<String>,
    #[serde(default)]
    title: Option<WasmChartExTitleInput>,
}

fn chart_ex_layout_from_string(value: String) -> duke_sheets_chart::ChartExLayout {
    match value.as_str() {
        "waterfall" => duke_sheets_chart::ChartExLayout::Waterfall,
        "treemap" => duke_sheets_chart::ChartExLayout::Treemap,
        "sunburst" => duke_sheets_chart::ChartExLayout::Sunburst,
        "funnel" => duke_sheets_chart::ChartExLayout::Funnel,
        "histogram" => duke_sheets_chart::ChartExLayout::Histogram,
        "boxWhisker" => duke_sheets_chart::ChartExLayout::BoxWhisker,
        "paretoLine" => duke_sheets_chart::ChartExLayout::ParetoLine,
        "regionMap" => duke_sheets_chart::ChartExLayout::RegionMap,
        "clusteredColumn" => duke_sheets_chart::ChartExLayout::ClusteredColumn,
        _ => duke_sheets_chart::ChartExLayout::Unknown(value),
    }
}

impl From<WasmChartExInput> for duke_sheets_chart::ChartEx {
    fn from(input: WasmChartExInput) -> Self {
        let title = input.title.map(|title| duke_sheets_chart::ChartExTitle {
            text: title.text,
            position: title.position,
            align: title.align,
            overlay: title.overlay,
            ..Default::default()
        });
        let series = duke_sheets_chart::ChartExSeries {
            layout: chart_ex_layout_from_string(input.layout),
            unique_id: None,
            hidden: None,
            owner_idx: None,
            format_idx: None,
            text: None,
            data_id: 0,
            data_labels: None,
            data_points: Vec::new(),
            layout_properties: None,
            axis_ids: Vec::new(),
            value_colors: None,
            value_color_positions: None,
            shape_properties: None,
            extensions: None,
        };
        Self {
            version: input.version,
            feature_list: input.feature_list,
            fallback_img: input.fallback_img,
            title,
            data: Vec::new(),
            external_data: None,
            plot_area: duke_sheets_chart::ChartExPlotArea {
                series: vec![series],
                ..Default::default()
            },
            legend: None,
            shape_properties: None,
            text_properties: None,
            color_map_override: None,
            format_overrides: Vec::new(),
            print_settings: None,
            raw_chart_style: None,
            raw_chart_color_style: None,
            extensions: None,
            raw_extensions: std::collections::HashMap::new(),
            raw_mc_fallback: None,
        }
    }
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmGroup {
    group_transform: WasmGroupTransform,
    children: Vec<WasmDrawing>,
}

#[derive(Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmGroupInput {
    #[serde(default)]
    group_transform: WasmGroupTransform,
    #[serde(default)]
    children: Vec<WasmDrawingInput>,
}

#[derive(Clone, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum WasmDrawingKind {
    Image { image: WasmImageMetadata },
    Chart { chart: WasmChart },
    ChartEx { chart_ex: WasmChartEx },
    FormControl { form_control: WasmFormControl },
    Comment { comment: WasmDrawingComment },
    Shape { shape: WasmShape },
    Group { group: WasmGroup },
    Raw { raw: WasmRawDrawingMetadata },
}

#[derive(Clone, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum WasmDrawingInputKind {
    Image { image: WasmImageInput },
    Chart { chart: WasmChartInput },
    ChartEx { chart_ex: WasmChartExInput },
    FormControl { form_control: WasmFormControl },
    Comment { comment: WasmDrawingComment },
    Shape { shape: WasmShape },
    Group { group: WasmGroupInput },
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmDrawing {
    drawing_path: Vec<usize>,
    /// Resolved on-sheet placement in EMU: the anchor rectangle for
    /// top-level drawings, the group-mapped (rotation/flip aware)
    /// rectangle for group children.
    absolute_rect_emu: WasmRectEmu,
    #[serde(flatten)]
    meta: WasmDrawingMeta,
    #[serde(flatten)]
    placement: WasmDrawingPlacement,
    #[serde(flatten)]
    kind: WasmDrawingKind,
}

#[derive(Clone, Copy, Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmRectEmu {
    x_emu: i64,
    y_emu: i64,
    width_emu: i64,
    height_emu: i64,
}

impl WasmRectEmu {
    fn from_core(rect: core::drawing::RectEmu) -> Self {
        // Core saturates at the JS safe-integer range, so the i64
        // values are directly representable.
        Self {
            x_emu: rect.x_emu,
            y_emu: rect.y_emu,
            width_emu: rect.width_emu,
            height_emu: rect.height_emu,
        }
    }
}

#[derive(Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WasmDrawingInput {
    #[serde(flatten)]
    meta: WasmDrawingMetaInput,
    #[serde(default)]
    anchor: Option<WasmDrawingAnchor>,
    #[serde(default)]
    transform: Option<WasmChildTransform>,
    #[serde(flatten)]
    kind: WasmDrawingInputKind,
}

impl WasmDrawing {
    fn from_object(sheet: &core::Worksheet, object: &core::DrawingObject, index: usize) -> Self {
        let path = vec![index];
        Self::from_parts(
            sheet,
            &object.meta,
            WasmDrawingPlacement::TopLevel {
                anchor: WasmDrawingAnchor::from(&object.anchor),
            },
            &object.kind,
            path,
        )
    }

    fn from_child(sheet: &core::Worksheet, child: &core::GroupChild, path: Vec<usize>) -> Self {
        Self::from_parts(
            sheet,
            &child.meta,
            WasmDrawingPlacement::Child {
                transform: WasmChildTransform::from(&child.transform),
            },
            &child.kind,
            path,
        )
    }

    fn from_parts(
        sheet: &core::Worksheet,
        meta: &core::DrawingMeta,
        placement: WasmDrawingPlacement,
        kind: &core::DrawingKind,
        drawing_path: Vec<usize>,
    ) -> Self {
        let kind = match kind {
            core::DrawingKind::Image(image) => WasmDrawingKind::Image {
                image: WasmImageMetadata::from(image),
            },
            core::DrawingKind::Chart(chart) => WasmDrawingKind::Chart {
                chart: WasmChart::from(chart.as_ref()),
            },
            core::DrawingKind::ChartEx(chart_ex) => WasmDrawingKind::ChartEx {
                chart_ex: WasmChartEx::from(chart_ex.as_ref()),
            },
            core::DrawingKind::FormControl(form_control) => WasmDrawingKind::FormControl {
                form_control: WasmFormControl::from(form_control),
            },
            core::DrawingKind::Comment { row, col, comment } => WasmDrawingKind::Comment {
                comment: WasmDrawingComment {
                    row: *row,
                    col: *col,
                    author: comment.author.clone(),
                    text: comment.text.clone(),
                },
            },
            core::DrawingKind::Shape(shape) => WasmDrawingKind::Shape {
                shape: WasmShape::from(shape.as_ref()),
            },
            core::DrawingKind::Group(group) => {
                let children = group
                    .children
                    .iter()
                    .enumerate()
                    .map(|(index, child)| {
                        let mut path = drawing_path.clone();
                        path.push(index);
                        Self::from_child(sheet, child, path)
                    })
                    .collect();
                WasmDrawingKind::Group {
                    group: WasmGroup {
                        group_transform: WasmGroupTransform::from(&group.transform),
                        children,
                    },
                }
            }
            core::DrawingKind::Raw(raw) => WasmDrawingKind::Raw {
                raw: WasmRawDrawingMetadata {
                    byte_length: raw.bytes.len(),
                    relationships: raw
                        .rels
                        .iter()
                        .map(|rel| WasmRawRelationshipMetadata {
                            id: rel.id.clone(),
                            rel_type: rel.rel_type.clone(),
                            target: rel.target.clone(),
                            external: rel.external,
                            has_part: rel.part.is_some(),
                        })
                        .collect(),
                },
            },
        };
        let absolute_rect_emu = WasmRectEmu::from_core(
            sheet
                .drawing_rect_emu(&drawing_path)
                .unwrap_or_default(),
        );
        Self {
            drawing_path,
            absolute_rect_emu,
            meta: WasmDrawingMeta::from(meta),
            placement,
            kind,
        }
    }

    fn flatten_into(&self, output: &mut Vec<Self>) {
        output.push(self.clone());
        if let WasmDrawingKind::Group { group } = &self.kind {
            for child in &group.children {
                child.flatten_into(output);
            }
        }
    }
}

impl WasmDrawingInput {
    fn into_object(self) -> Result<core::DrawingObject, String> {
        let anchor = self
            .anchor
            .ok_or_else(|| "top-level drawing input requires anchor".to_string())?;
        if self.transform.is_some() {
            return Err("top-level drawing input cannot contain transform".to_string());
        }
        let (kind, comment) = drawing_kind_from_input(self.kind)?;
        let object = core::DrawingObject {
            meta: self.meta.into_core(comment),
            anchor: anchor.try_into()?,
            kind,
        };
        object.validate().map_err(|error| error.to_string())?;
        Ok(object)
    }

    fn into_child(self) -> Result<core::GroupChild, String> {
        let transform = self
            .transform
            .ok_or_else(|| "group child input requires transform".to_string())?;
        if self.anchor.is_some() {
            return Err("group child input cannot contain anchor".to_string());
        }
        let (kind, comment) = drawing_kind_from_input(self.kind)?;
        let child = core::GroupChild {
            meta: self.meta.into_core(comment),
            transform: transform.into(),
            kind,
        };
        // Validate in group-child position so kinds that cannot nest
        // (comments, raw) are rejected here, not at write time.
        core::DrawingObject::group(core::Group {
            transform: GroupTransform::default(),
            children: vec![child.clone()],
        })
        .validate()
        .map_err(|error| error.to_string())?;
        Ok(child)
    }
}

fn drawing_kind_from_input(kind: WasmDrawingInputKind) -> Result<(core::DrawingKind, bool), String> {
    Ok(match kind {
        WasmDrawingInputKind::Image { image } => {
            (core::DrawingKind::Image(image.try_into()?), false)
        }
        WasmDrawingInputKind::Chart { chart } => {
            (core::DrawingKind::Chart(Box::new(chart.try_into()?)), false)
        }
        WasmDrawingInputKind::ChartEx { chart_ex } => (
            core::DrawingKind::ChartEx(Box::new(chart_ex.into())),
            false,
        ),
        WasmDrawingInputKind::FormControl { form_control } => (
            core::DrawingKind::FormControl(form_control.try_into()?),
            false,
        ),
        WasmDrawingInputKind::Comment { comment } => (
            core::DrawingKind::Comment {
                row: comment.row,
                col: comment.col,
                comment: core::CellComment::new(comment.author, comment.text),
            },
            true,
        ),
        WasmDrawingInputKind::Shape { shape } => {
            (core::DrawingKind::Shape(Box::new(shape.try_into()?)), false)
        }
        WasmDrawingInputKind::Group { group } => {
            let children = group
                .children
                .into_iter()
                .map(WasmDrawingInput::into_child)
                .collect::<Result<Vec<_>, _>>()?;
            (
                core::DrawingKind::Group(Box::new(core::Group {
                    transform: group.group_transform.into(),
                    children,
                })),
                false,
            )
        }
    })
}

fn deserialize_drawing(value: JsValue) -> Result<WasmDrawingInput, JsError> {
    serde_wasm_bindgen::from_value(value)
        .map_err(|error| JsError::new(&format!("invalid drawing: {error}")))
}

fn deserialize_path(value: JsValue) -> Result<Vec<usize>, JsError> {
    let path: Vec<usize> = serde_wasm_bindgen::from_value(value)
        .map_err(|error| JsError::new(&format!("invalid drawing path: {error}")))?;
    if path.is_empty() {
        return Err(JsError::new("drawing path cannot be empty"));
    }
    Ok(path)
}

fn drawing_tree(sheet: &core::Worksheet) -> Vec<WasmDrawing> {
    sheet
        .drawings()
        .iter()
        .enumerate()
        .map(|(index, object)| WasmDrawing::from_object(sheet, object, index))
        .collect()
}

fn drawings_flat(sheet: &core::Worksheet) -> Vec<WasmDrawing> {
    let mut flat = Vec::new();
    for drawing in drawing_tree(sheet) {
        drawing.flatten_into(&mut flat);
    }
    flat
}

#[wasm_bindgen]
impl Worksheet {
    /// The recursive drawing tree in top-level z-order.
    #[wasm_bindgen(getter, js_name = drawings, skip_typescript)]
    pub fn drawings(&self) -> Result<JsValue, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        to_js_value(&drawing_tree(sheet))
    }

    /// All form controls in depth-first drawing order.
    #[wasm_bindgen(getter, js_name = formControls, skip_typescript)]
    pub fn form_controls(&self) -> Result<JsValue, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let controls: Vec<_> = drawings_flat(sheet)
            .into_iter()
            .filter(|drawing| matches!(drawing.kind, WasmDrawingKind::FormControl { .. }))
            .collect();
        to_js_value(&controls)
    }

    #[wasm_bindgen(getter, js_name = formControlCount, skip_typescript)]
    pub fn form_control_count(&self) -> Result<u32, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        u32::try_from(sheet.placed_form_controls().len())
            .map_err(|_| JsError::new("form control count exceeds u32"))
    }

    /// All embedded images in depth-first drawing order, without image bytes.
    #[wasm_bindgen(getter, js_name = images, skip_typescript)]
    pub fn images(&self) -> Result<JsValue, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let images: Vec<_> = drawings_flat(sheet)
            .into_iter()
            .filter(|drawing| matches!(drawing.kind, WasmDrawingKind::Image { .. }))
            .collect();
        to_js_value(&images)
    }

    #[wasm_bindgen(getter, js_name = imageCount, skip_typescript)]
    pub fn image_count(&self) -> Result<u32, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let count = sheet
            .drawings_flat()
            .into_iter()
            .filter(|(_, node)| matches!(node.kind, core::DrawingKind::Image(_)))
            .count();
        u32::try_from(count).map_err(|_| JsError::new("image count exceeds u32"))
    }

    /// All standard charts in depth-first drawing order.
    #[wasm_bindgen(getter, js_name = charts, skip_typescript)]
    pub fn charts(&self) -> Result<JsValue, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let charts: Vec<_> = drawings_flat(sheet)
            .into_iter()
            .filter(|drawing| matches!(drawing.kind, WasmDrawingKind::Chart { .. }))
            .collect();
        to_js_value(&charts)
    }

    #[wasm_bindgen(getter, js_name = chartCount, skip_typescript)]
    pub fn chart_count(&self) -> Result<u32, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let count = sheet
            .drawings_flat()
            .into_iter()
            .filter(|(_, node)| matches!(node.kind, core::DrawingKind::Chart(_)))
            .count();
        u32::try_from(count).map_err(|_| JsError::new("chart count exceeds u32"))
    }

    /// All ChartEx charts in depth-first drawing order.
    #[wasm_bindgen(getter, js_name = chartsEx, skip_typescript)]
    pub fn charts_ex(&self) -> Result<JsValue, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let charts: Vec<_> = drawings_flat(sheet)
            .into_iter()
            .filter(|drawing| matches!(drawing.kind, WasmDrawingKind::ChartEx { .. }))
            .collect();
        to_js_value(&charts)
    }

    #[wasm_bindgen(getter, js_name = chartExCount, skip_typescript)]
    pub fn chart_ex_count(&self) -> Result<u32, JsError> {
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let count = sheet
            .drawings_flat()
            .into_iter()
            .filter(|(_, node)| matches!(node.kind, core::DrawingKind::ChartEx(_)))
            .count();
        u32::try_from(count).map_err(|_| JsError::new("ChartEx count exceeds u32"))
    }

    /// Append a top-level drawing and return its z-order index.
    #[wasm_bindgen(js_name = addDrawing, skip_typescript)]
    pub fn add_drawing(&self, input: JsValue) -> Result<u32, JsError> {
        let object = deserialize_drawing(input)?
            .into_object()
            .map_err(|error| JsError::new(&error))?;
        let mut workbook = self.workbook.borrow_mut();
        let sheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let index = sheet.add_drawing(object).map_err(to_js_error)?;
        u32::try_from(index).map_err(|_| JsError::new("drawing index exceeds u32"))
    }

    /// Insert a top-level drawing at a z-order index.
    #[wasm_bindgen(js_name = insertDrawing, skip_typescript)]
    pub fn insert_drawing(&self, index: u32, input: JsValue) -> Result<(), JsError> {
        let object = deserialize_drawing(input)?
            .into_object()
            .map_err(|error| JsError::new(&error))?;
        let mut workbook = self.workbook.borrow_mut();
        let sheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        sheet
            .insert_drawing(index as usize, object)
            .map_err(to_js_error)
    }

    /// Replace a top-level drawing or nested group child.
    #[wasm_bindgen(js_name = setDrawing, skip_typescript)]
    pub fn set_drawing(&self, path: JsValue, input: JsValue) -> Result<(), JsError> {
        let path = deserialize_path(path)?;
        let input = deserialize_drawing(input)?;
        let mut workbook = self.workbook.borrow_mut();
        let sheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        if path.len() == 1 {
            let object = input
                .into_object()
                .map_err(|error| JsError::new(&error))?;
            return sheet.set_drawing(path[0], object).map_err(to_js_error);
        }

        let child = input
            .into_child()
            .map_err(|error| JsError::new(&error))?;
        sheet.set_group_child(&path, child).map_err(to_js_error)
    }

    /// Remove a top-level drawing or nested group child.
    #[wasm_bindgen(js_name = removeDrawing, skip_typescript)]
    pub fn remove_drawing(&self, path: JsValue) -> Result<(), JsError> {
        let path = deserialize_path(path)?;
        let mut workbook = self.workbook.borrow_mut();
        let sheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        if path.len() == 1 {
            return sheet
                .remove_drawing(path[0])
                .map(|_| ())
                .map_err(to_js_error);
        }

        sheet
            .remove_group_child(&path)
            .map(|_| ())
            .map_err(to_js_error)
    }

    /// Move a top-level drawing to another z-order index.
    #[wasm_bindgen(js_name = moveDrawing, skip_typescript)]
    pub fn move_drawing(&self, from: u32, to: u32) -> Result<(), JsError> {
        let mut workbook = self.workbook.borrow_mut();
        let sheet = workbook
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        sheet
            .move_drawing(from as usize, to as usize)
            .map_err(to_js_error)
    }

    /// Return image bytes for an image drawing without copying them into metadata getters.
    #[wasm_bindgen(js_name = drawingImageData, skip_typescript)]
    pub fn drawing_image_data(&self, path: JsValue) -> Result<Vec<u8>, JsError> {
        let path = deserialize_path(path)?;
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let node = sheet
            .drawing_at_path(&path)
            .ok_or_else(|| JsError::new(&format!("drawing path {path:?} not found")))?;
        let core::DrawingKind::Image(image) = node.kind else {
            return Err(JsError::new("drawing path does not identify an image"));
        };
        Ok(image.data().to_vec())
    }

    /// Return an image drawing's SVG companion bytes, when present.
    #[wasm_bindgen(js_name = drawingSvgData, skip_typescript)]
    pub fn drawing_svg_data(&self, path: JsValue) -> Result<Option<Vec<u8>>, JsError> {
        let path = deserialize_path(path)?;
        let workbook = self.workbook.borrow();
        let sheet = workbook
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let node = sheet
            .drawing_at_path(&path)
            .ok_or_else(|| JsError::new(&format!("drawing path {path:?} not found")))?;
        let core::DrawingKind::Image(image) = node.kind else {
            return Err(JsError::new("drawing path does not identify an image"));
        };
        Ok(image.svg_data().map(ToOwned::to_owned))
    }

    /// Apply Excel checkbox/radio semantics and synchronize linked cells immediately.
    #[wasm_bindgen(js_name = setFormControlCheckState, skip_typescript)]
    pub fn set_form_control_check_state(
        &self,
        path: JsValue,
        state: &str,
    ) -> Result<JsValue, JsError> {
        let path = deserialize_path(path)?;
        let state = match state {
            "unchecked" => core::CheckState::Unchecked,
            "checked" => core::CheckState::Checked,
            "mixed" => core::CheckState::Mixed,
            other => return Err(JsError::new(&format!("invalid check state {other:?}"))),
        };
        let mut workbook = self.workbook.borrow_mut();
        let result = workbook
            .set_form_control_check_state(self.sheet_index, &path, state)
            .map_err(to_js_error)?;
        to_js_value(&WasmFormControlInteractionResult {
            controls_changed: result.controls_changed,
            linked_cells_changed: result.linked_cells_changed,
        })
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmFormControlInteractionResult {
    controls_changed: usize,
    linked_cells_changed: usize,
}

#[wasm_bindgen]
impl Workbook {
    /// Project all form-control state into linked cells.
    #[wasm_bindgen(js_name = syncFormControls)]
    pub fn sync_form_controls(&self) -> usize {
        self.inner.borrow_mut().sync_form_control_links()
    }

    /// Drive controls from formula-backed linked cells.
    #[wasm_bindgen(js_name = syncFormControlsFromLinkedCells)]
    pub fn sync_form_controls_from_linked_cells(&self) -> usize {
        self.inner
            .borrow_mut()
            .sync_form_controls_from_linked_cells()
    }
}
