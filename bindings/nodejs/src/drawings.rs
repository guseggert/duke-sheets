use napi::bindgen_prelude::*;
use napi_derive::napi;
use serde::{Deserialize, Serialize};
use serde_bytes::ByteBuf;

use duke_sheets_chart::{
    CellMarker, ChildTransform, DrawingAnchor, EditAs, EmbeddedImage, GroupTransform, ImageFormat,
};
use duke_sheets_core as core;

use super::{catch_panic, to_napi_err, JsChart, JsChartEx, JsCheckState, Workbook, Worksheet};

type JsObject<'env> = Object<'env>;

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingMetaInput {
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

impl DrawingMetaInput {
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
struct DrawingCellMarker {
    col: u16,
    row: u32,
    #[serde(default)]
    col_offset_emu: i64,
    #[serde(default)]
    row_offset_emu: i64,
}

impl From<&CellMarker> for DrawingCellMarker {
    fn from(marker: &CellMarker) -> Self {
        Self {
            col: marker.col,
            row: marker.row,
            col_offset_emu: marker.col_offset_emu,
            row_offset_emu: marker.row_offset_emu,
        }
    }
}

impl From<DrawingCellMarker> for CellMarker {
    fn from(marker: DrawingCellMarker) -> Self {
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
enum DrawingAnchorDto {
    TwoCell {
        from: DrawingCellMarker,
        to: DrawingCellMarker,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        edit_as: Option<String>,
    },
    OneCell {
        from: DrawingCellMarker,
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

impl From<&DrawingAnchor> for DrawingAnchorDto {
    fn from(anchor: &DrawingAnchor) -> Self {
        match anchor {
            DrawingAnchor::TwoCell { from, to, edit_as } => Self::TwoCell {
                from: DrawingCellMarker::from(from),
                to: DrawingCellMarker::from(to),
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
                from: DrawingCellMarker::from(from),
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

impl TryFrom<DrawingAnchorDto> for DrawingAnchor {
    type Error = String;

    fn try_from(anchor: DrawingAnchorDto) -> std::result::Result<Self, Self::Error> {
        Ok(match anchor {
            DrawingAnchorDto::TwoCell { from, to, edit_as } => DrawingAnchor::TwoCell {
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
            DrawingAnchorDto::OneCell {
                from,
                width_emu,
                height_emu,
            } => DrawingAnchor::OneCell {
                from: from.into(),
                width_emu,
                height_emu,
            },
            DrawingAnchorDto::Absolute {
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

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct DrawingChildTransform {
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

impl From<&ChildTransform> for DrawingChildTransform {
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

impl From<DrawingChildTransform> for ChildTransform {
    fn from(transform: DrawingChildTransform) -> Self {
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
struct DrawingGroupTransform {
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

impl From<&GroupTransform> for DrawingGroupTransform {
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

impl From<DrawingGroupTransform> for GroupTransform {
    fn from(transform: DrawingGroupTransform) -> Self {
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

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "colorType", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum DrawingColor {
    Auto,
    Rgb { r: u8, g: u8, b: u8 },
    Argb { a: u8, r: u8, g: u8, b: u8 },
    Theme { index: u8, tint: i8 },
    Indexed { index: u8 },
}

impl From<&core::Color> for DrawingColor {
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

impl From<DrawingColor> for core::Color {
    fn from(color: DrawingColor) -> Self {
        match color {
            DrawingColor::Auto => Self::Auto,
            DrawingColor::Rgb { r, g, b } => Self::Rgb { r, g, b },
            DrawingColor::Argb { a, r, g, b } => Self::Argb { a, r, g, b },
            DrawingColor::Theme { index, tint } => Self::Theme { index, tint },
            DrawingColor::Indexed { index } => Self::Indexed(index),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct DrawingRunFont {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    color: Option<DrawingColor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    underline: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    strikethrough: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    vertical_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    family: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    charset: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    scheme: Option<String>,
}

impl From<&core::RunFont> for DrawingRunFont {
    fn from(font: &core::RunFont) -> Self {
        Self {
            bold: font.bold,
            italic: font.italic,
            size: font.size,
            color: font.color.as_ref().map(DrawingColor::from),
            name: font.name.clone(),
            underline: font.underline.map(|underline| match underline {
                core::style::Underline::None => "none".to_string(),
                core::style::Underline::Single => "single".to_string(),
                core::style::Underline::Double => "double".to_string(),
                core::style::Underline::SingleAccounting => "singleAccounting".to_string(),
                core::style::Underline::DoubleAccounting => "doubleAccounting".to_string(),
            }),
            strikethrough: font.strikethrough,
            vertical_align: font.vertical_align.map(|alignment| match alignment {
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

impl TryFrom<DrawingRunFont> for core::RunFont {
    type Error = String;

    fn try_from(font: DrawingRunFont) -> std::result::Result<Self, Self::Error> {
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
                .map(|alignment| match alignment.as_str() {
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
struct DrawingTextRun {
    text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    font: Option<DrawingRunFont>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct DrawingText {
    #[serde(default)]
    runs: Vec<DrawingTextRun>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    horizontal_alignment: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    vertical_alignment: Option<String>,
}

impl From<&core::DrawingText> for DrawingText {
    fn from(text: &core::DrawingText) -> Self {
        Self {
            runs: text
                .runs
                .iter()
                .map(|run| DrawingTextRun {
                    text: run.text.clone(),
                    font: run.font.as_ref().map(DrawingRunFont::from),
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

impl TryFrom<DrawingText> for core::DrawingText {
    type Error = String;

    fn try_from(text: DrawingText) -> std::result::Result<Self, Self::Error> {
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
            .collect::<std::result::Result<Vec<_>, String>>()?;
        Ok(Self {
            runs,
            horizontal_alignment,
            vertical_alignment,
        })
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
enum DrawingCheckState {
    Unchecked,
    Checked,
    Mixed,
}

impl From<core::CheckState> for DrawingCheckState {
    fn from(state: core::CheckState) -> Self {
        match state {
            core::CheckState::Unchecked => Self::Unchecked,
            core::CheckState::Checked => Self::Checked,
            core::CheckState::Mixed => Self::Mixed,
        }
    }
}

impl From<DrawingCheckState> for core::CheckState {
    fn from(state: DrawingCheckState) -> Self {
        match state {
            DrawingCheckState::Unchecked => Self::Unchecked,
            DrawingCheckState::Checked => Self::Checked,
            DrawingCheckState::Mixed => Self::Mixed,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
enum DrawingListSelection {
    Single,
    Multi,
    Extend,
}

impl From<core::ListSelection> for DrawingListSelection {
    fn from(selection: core::ListSelection) -> Self {
        match selection {
            core::ListSelection::Single => Self::Single,
            core::ListSelection::Multi => Self::Multi,
            core::ListSelection::Extend => Self::Extend,
        }
    }
}

impl From<DrawingListSelection> for core::ListSelection {
    fn from(selection: DrawingListSelection) -> Self {
        match selection {
            DrawingListSelection::Single => Self::Single,
            DrawingListSelection::Multi => Self::Multi,
            DrawingListSelection::Extend => Self::Extend,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum DrawingFormControlKind {
    Button {
        caption: DrawingText,
    },
    Checkbox {
        caption: DrawingText,
        state: DrawingCheckState,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        cell_link: Option<String>,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    OptionButton {
        caption: DrawingText,
        state: DrawingCheckState,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        cell_link: Option<String>,
        #[serde(default)]
        first_in_group: bool,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    Label {
        caption: DrawingText,
    },
    GroupBox {
        caption: DrawingText,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    ListBox {
        #[serde(default, skip_serializing_if = "Option::is_none")]
        input_range: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        cell_link: Option<String>,
        selection: DrawingListSelection,
        #[serde(default)]
        selected: Vec<u16>,
        #[serde(rename = "no3D", default)]
        no_3d: bool,
    },
    Dropdown {
        #[serde(default, skip_serializing_if = "Option::is_none")]
        input_range: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        cell_link: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
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
        #[serde(default, skip_serializing_if = "Option::is_none")]
        cell_link: Option<String>,
    },
    Spinner {
        value: u16,
        min: u16,
        max: u16,
        increment: u16,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        cell_link: Option<String>,
    },
    Unknown {
        object_type: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        legacy_object_type: Option<u16>,
        #[serde(default)]
        caption: DrawingText,
        /// Internal passthrough of unmodeled XLSX `formControlPr`
        /// attributes; echoed back unchanged on write.
        #[serde(default)]
        raw_properties: Vec<(String, String)>,
        /// Internal passthrough of the original BIFF OBJ body (Buffer
        /// in JS), required for XLS rewrite.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        raw_obj: Option<ByteBuf>,
    },
}

impl From<&core::FormControlKind> for DrawingFormControlKind {
    fn from(kind: &core::FormControlKind) -> Self {
        match kind {
            core::FormControlKind::Button { caption } => Self::Button {
                caption: DrawingText::from(caption),
            },
            core::FormControlKind::Checkbox {
                caption,
                state,
                cell_link,
                no_3d,
            } => Self::Checkbox {
                caption: DrawingText::from(caption),
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
                caption: DrawingText::from(caption),
                state: (*state).into(),
                cell_link: cell_link.clone(),
                first_in_group: *first_in_group,
                no_3d: *no_3d,
            },
            core::FormControlKind::Label { caption } => Self::Label {
                caption: DrawingText::from(caption),
            },
            core::FormControlKind::GroupBox { caption, no_3d } => Self::GroupBox {
                caption: DrawingText::from(caption),
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
                caption: DrawingText::from(caption),
                raw_properties: raw_properties.clone(),
                raw_obj: raw_obj.clone().map(ByteBuf::from),
            },
        }
    }
}

impl TryFrom<DrawingFormControlKind> for core::FormControlKind {
    type Error = String;

    fn try_from(kind: DrawingFormControlKind) -> std::result::Result<Self, Self::Error> {
        Ok(match kind {
            DrawingFormControlKind::Button { caption } => Self::Button {
                caption: caption.try_into()?,
            },
            DrawingFormControlKind::Checkbox {
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
            DrawingFormControlKind::OptionButton {
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
            DrawingFormControlKind::Label { caption } => Self::Label {
                caption: caption.try_into()?,
            },
            DrawingFormControlKind::GroupBox { caption, no_3d } => Self::GroupBox {
                caption: caption.try_into()?,
                no_3d,
            },
            DrawingFormControlKind::ListBox {
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
            DrawingFormControlKind::Dropdown {
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
            DrawingFormControlKind::Scrollbar {
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
            DrawingFormControlKind::Spinner {
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
            DrawingFormControlKind::Unknown {
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
                raw_obj: raw_obj.map(ByteBuf::into_vec),
            },
        })
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingFormControl {
    kind: DrawingFormControlKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    macro_name: Option<String>,
    /// Unmodeled VML `ClientData` children (Buffers in JS) preserved
    /// on any control kind; echoed back unchanged on write.
    #[serde(default)]
    raw_client_data: Vec<ByteBuf>,
}

impl From<&core::FormControl> for DrawingFormControl {
    fn from(control: &core::FormControl) -> Self {
        Self {
            kind: DrawingFormControlKind::from(&control.kind),
            macro_name: control.macro_name.clone(),
            raw_client_data: control
                .raw_client_data
                .iter()
                .cloned()
                .map(ByteBuf::from)
                .collect(),
        }
    }
}

impl TryFrom<DrawingFormControl> for core::FormControl {
    type Error = String;

    fn try_from(control: DrawingFormControl) -> std::result::Result<Self, Self::Error> {
        let result = Self {
            kind: control.kind.try_into()?,
            macro_name: control.macro_name,
            raw_client_data: control
                .raw_client_data
                .into_iter()
                .map(ByteBuf::into_vec)
                .collect(),
        };
        result.validate().map_err(|error| error.to_string())?;
        Ok(result)
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct DrawingImageMetadata {
    format: String,
    media_path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    svg_media_path: Option<String>,
    width_emu: i64,
    height_emu: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    rotation: Option<i32>,
    flip_h: bool,
    flip_v: bool,
}

impl From<&EmbeddedImage> for DrawingImageMetadata {
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

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingImageInput {
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
    data: ByteBuf,
    #[serde(default)]
    svg_data: Option<ByteBuf>,
}

impl TryFrom<DrawingImageInput> for EmbeddedImage {
    type Error = String;

    fn try_from(image: DrawingImageInput) -> std::result::Result<Self, Self::Error> {
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
            data: image.data.into_vec(),
            svg_data: image.svg_data.map(ByteBuf::into_vec),
        })
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum DrawingShapeFill {
    None,
    Solid { color: DrawingColor },
}

impl From<&core::ShapeFill> for DrawingShapeFill {
    fn from(fill: &core::ShapeFill) -> Self {
        match fill {
            core::ShapeFill::None => Self::None,
            core::ShapeFill::Solid(color) => Self::Solid {
                color: DrawingColor::from(color),
            },
        }
    }
}

impl From<DrawingShapeFill> for core::ShapeFill {
    fn from(fill: DrawingShapeFill) -> Self {
        match fill {
            DrawingShapeFill::None => Self::None,
            DrawingShapeFill::Solid { color } => Self::Solid(color.into()),
        }
    }
}

fn default_shape_fill() -> DrawingShapeFill {
    DrawingShapeFill::None
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct DrawingShapeLine {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    color: Option<DrawingColor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    width_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    dash_style: Option<String>,
    #[serde(default)]
    no_fill: bool,
}

impl From<&core::ShapeLine> for DrawingShapeLine {
    fn from(line: &core::ShapeLine) -> Self {
        Self {
            color: line.color.as_ref().map(DrawingColor::from),
            width_emu: line.width_emu,
            dash_style: line.dash_style.clone(),
            no_fill: line.no_fill,
        }
    }
}

impl From<DrawingShapeLine> for core::ShapeLine {
    fn from(line: DrawingShapeLine) -> Self {
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

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingShape {
    #[serde(default = "default_shape_geometry")]
    geometry: String,
    #[serde(default = "default_shape_fill")]
    fill: DrawingShapeFill,
    #[serde(default)]
    line: DrawingShapeLine,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    text: Option<DrawingText>,
    #[serde(default)]
    rotation: i32,
    #[serde(default)]
    flip_h: bool,
    #[serde(default)]
    flip_v: bool,
}

impl From<&core::Shape> for DrawingShape {
    fn from(shape: &core::Shape) -> Self {
        let core::ShapeGeometry::Preset(geometry) = &shape.geometry;
        Self {
            geometry: geometry.clone(),
            fill: DrawingShapeFill::from(&shape.fill),
            line: DrawingShapeLine::from(&shape.line),
            text: shape.text.as_ref().map(DrawingText::from),
            rotation: shape.rotation,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
        }
    }
}

impl TryFrom<DrawingShape> for core::Shape {
    type Error = String;

    fn try_from(shape: DrawingShape) -> std::result::Result<Self, Self::Error> {
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
struct DrawingComment {
    row: u32,
    col: u16,
    author: String,
    text: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(tag = "refType", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum DrawingChartDataReferenceInput {
    Formula { formula: String },
    Numbers { numbers: Vec<f64> },
    Strings { strings: Vec<String> },
}

impl From<DrawingChartDataReferenceInput> for duke_sheets_chart::DataReference {
    fn from(reference: DrawingChartDataReferenceInput) -> Self {
        match reference {
            DrawingChartDataReferenceInput::Formula { formula } => Self::Formula(formula),
            DrawingChartDataReferenceInput::Numbers { numbers } => Self::Numbers(numbers),
            DrawingChartDataReferenceInput::Strings { strings } => Self::Strings(strings),
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingChartSeriesInput {
    #[serde(default)]
    name: Option<String>,
    values: DrawingChartDataReferenceInput,
    #[serde(default)]
    categories: Option<DrawingChartDataReferenceInput>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingChartInput {
    chart_type: String,
    #[serde(default)]
    title: Option<String>,
    #[serde(default)]
    series: Vec<DrawingChartSeriesInput>,
    #[serde(rename = "is3D", default)]
    is_3d: bool,
    #[serde(default)]
    vary_colors: Option<bool>,
    #[serde(default)]
    gap_width: Option<u32>,
    #[serde(default)]
    overlap: Option<i32>,
}

fn chart_type_from_string(
    value: &str,
) -> std::result::Result<duke_sheets_chart::ChartType, String> {
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

impl TryFrom<DrawingChartInput> for duke_sheets_chart::Chart {
    type Error = String;

    fn try_from(input: DrawingChartInput) -> std::result::Result<Self, Self::Error> {
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

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingChartExTitleInput {
    #[serde(default)]
    text: Option<String>,
    #[serde(default)]
    position: Option<String>,
    #[serde(default)]
    align: Option<String>,
    #[serde(default)]
    overlay: Option<bool>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingChartExInput {
    layout: String,
    #[serde(default)]
    version: Option<String>,
    #[serde(default)]
    feature_list: Option<String>,
    #[serde(default)]
    fallback_img: Option<String>,
    #[serde(default)]
    title: Option<DrawingChartExTitleInput>,
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

impl From<DrawingChartExInput> for duke_sheets_chart::ChartEx {
    fn from(input: DrawingChartExInput) -> Self {
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

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingGroupInput {
    #[serde(default)]
    group_transform: DrawingGroupTransform,
    #[serde(default)]
    children: Vec<DrawingInput>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
enum DrawingInputKind {
    Image { image: DrawingImageInput },
    Chart { chart: DrawingChartInput },
    ChartEx { chart_ex: DrawingChartExInput },
    FormControl { form_control: DrawingFormControl },
    Comment { comment: DrawingComment },
    Shape { shape: DrawingShape },
    Group { group: DrawingGroupInput },
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DrawingInput {
    #[serde(flatten)]
    meta: DrawingMetaInput,
    #[serde(default)]
    anchor: Option<DrawingAnchorDto>,
    #[serde(default)]
    transform: Option<DrawingChildTransform>,
    #[serde(flatten)]
    kind: DrawingInputKind,
}

impl DrawingInput {
    fn into_object(self) -> std::result::Result<core::DrawingObject, String> {
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

    fn into_child(self) -> std::result::Result<core::GroupChild, String> {
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

fn drawing_kind_from_input(
    kind: DrawingInputKind,
) -> std::result::Result<(core::DrawingKind, bool), String> {
    Ok(match kind {
        DrawingInputKind::Image { image } => {
            (core::DrawingKind::Image(image.try_into()?), false)
        }
        DrawingInputKind::Chart { chart } => {
            (core::DrawingKind::Chart(Box::new(chart.try_into()?)), false)
        }
        DrawingInputKind::ChartEx { chart_ex } => (
            core::DrawingKind::ChartEx(Box::new(chart_ex.into())),
            false,
        ),
        DrawingInputKind::FormControl { form_control } => (
            core::DrawingKind::FormControl(form_control.try_into()?),
            false,
        ),
        DrawingInputKind::Comment { comment } => (
            core::DrawingKind::Comment {
                row: comment.row,
                col: comment.col,
                comment: core::CellComment::new(comment.author, comment.text),
            },
            true,
        ),
        DrawingInputKind::Shape { shape } => {
            (core::DrawingKind::Shape(Box::new(shape.try_into()?)), false)
        }
        DrawingInputKind::Group { group } => {
            let children = group
                .children
                .into_iter()
                .map(DrawingInput::into_child)
                .collect::<std::result::Result<Vec<_>, _>>()?;
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct RawDrawingRelationshipMetadata {
    id: String,
    rel_type: String,
    target: String,
    external: bool,
    has_part: bool,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct RawDrawingMetadata {
    byte_length: usize,
    relationships: Vec<RawDrawingRelationshipMetadata>,
}

enum DrawingPlacement<'a> {
    TopLevel(&'a DrawingAnchor),
    Child(&'a ChildTransform),
}

fn set_serialized<'env, T: Serialize>(
    env: &'env Env,
    object: &mut JsObject<'env>,
    key: &str,
    value: &T,
) -> Result<()> {
    object.set(key, env.to_js_value(value)?)
}

fn drawing_path_to_js(path: &[usize]) -> Result<Vec<u32>> {
    path.iter()
        .map(|&index| {
            u32::try_from(index)
                .map_err(|_| napi::Error::from_reason("drawing path index exceeds u32"))
        })
        .collect()
}

fn objects_to_array<'env>(
    env: &'env Env,
    objects: Vec<JsObject<'env>>,
) -> Result<JsObject<'env>> {
    let mut array = env.create_array(objects.len() as u32)?;
    for (index, object) in objects.into_iter().enumerate() {
        array.set(index as u32, object)?;
    }
    array.coerce_to_object()
}

fn drawing_node_to_js<'env>(
    env: &'env Env,
    sheet: &core::Worksheet,
    meta: &core::DrawingMeta,
    placement: DrawingPlacement<'_>,
    kind: &core::DrawingKind,
    path: &[usize],
) -> Result<JsObject<'env>> {
    let mut drawing = Object::new(env)?;
    drawing.set("drawingPath", drawing_path_to_js(path)?)?;
    drawing.set(
        "absoluteRectEmu",
        rect_emu_to_js(env, sheet.drawing_rect_emu(path).unwrap_or_default())?,
    )?;
    if let Some(name) = &meta.name {
        drawing.set("name", name.clone())?;
    }
    drawing.set("hidden", meta.hidden)?;
    drawing.set("locked", meta.locked)?;
    drawing.set("printable", meta.printable)?;
    if let Some(alt_text) = &meta.alt_text {
        drawing.set("altText", alt_text.clone())?;
    }
    if let Some(title) = &meta.title {
        drawing.set("title", title.clone())?;
    }
    match placement {
        DrawingPlacement::TopLevel(anchor) => {
            set_serialized(env, &mut drawing, "anchor", &DrawingAnchorDto::from(anchor))?;
        }
        DrawingPlacement::Child(transform) => {
            set_serialized(
                env,
                &mut drawing,
                "transform",
                &DrawingChildTransform::from(transform),
            )?;
        }
    }

    match kind {
        core::DrawingKind::Image(image) => {
            drawing.set("kind", "image")?;
            set_serialized(
                env,
                &mut drawing,
                "image",
                &DrawingImageMetadata::from(image),
            )?;
        }
        core::DrawingKind::Chart(chart) => {
            drawing.set("kind", "chart")?;
            drawing.set("chart", JsChart::from(chart.as_ref()))?;
        }
        core::DrawingKind::ChartEx(chart_ex) => {
            drawing.set("kind", "chartEx")?;
            drawing.set("chartEx", JsChartEx::from(chart_ex.as_ref()))?;
        }
        core::DrawingKind::FormControl(form_control) => {
            drawing.set("kind", "formControl")?;
            set_serialized(
                env,
                &mut drawing,
                "formControl",
                &DrawingFormControl::from(form_control),
            )?;
        }
        core::DrawingKind::Comment { row, col, comment } => {
            drawing.set("kind", "comment")?;
            set_serialized(
                env,
                &mut drawing,
                "comment",
                &DrawingComment {
                    row: *row,
                    col: *col,
                    author: comment.author.clone(),
                    text: comment.text.clone(),
                },
            )?;
        }
        core::DrawingKind::Shape(shape) => {
            drawing.set("kind", "shape")?;
            set_serialized(
                env,
                &mut drawing,
                "shape",
                &DrawingShape::from(shape.as_ref()),
            )?;
        }
        core::DrawingKind::Group(group) => {
            drawing.set("kind", "group")?;
            let mut group_object = Object::new(env)?;
            set_serialized(
                env,
                &mut group_object,
                "groupTransform",
                &DrawingGroupTransform::from(&group.transform),
            )?;
            let mut children = Vec::with_capacity(group.children.len());
            for (index, child) in group.children.iter().enumerate() {
                let mut child_path = path.to_vec();
                child_path.push(index);
                children.push(drawing_node_to_js(
                    env,
                    sheet,
                    &child.meta,
                    DrawingPlacement::Child(&child.transform),
                    &child.kind,
                    &child_path,
                )?);
            }
            group_object.set("children", objects_to_array(env, children)?)?;
            drawing.set("group", group_object)?;
        }
        core::DrawingKind::Raw(raw) => {
            drawing.set("kind", "raw")?;
            set_serialized(
                env,
                &mut drawing,
                "raw",
                &RawDrawingMetadata {
                    byte_length: raw.bytes.len(),
                    relationships: raw
                        .rels
                        .iter()
                        .map(|rel| RawDrawingRelationshipMetadata {
                            id: rel.id.clone(),
                            rel_type: rel.rel_type.clone(),
                            target: rel.target.clone(),
                            external: rel.external,
                            has_part: rel.part.is_some(),
                        })
                        .collect(),
                },
            )?;
        }
    }
    Ok(drawing)
}

/// Resolved on-sheet placement in EMU. Core saturates the values at
/// the JS safe-integer range, so the `as f64` casts are exact.
fn rect_emu_to_js<'env>(
    env: &'env Env,
    rect: core::drawing::RectEmu,
) -> Result<JsObject<'env>> {
    let mut object = Object::new(env)?;
    object.set("xEmu", rect.x_emu as f64)?;
    object.set("yEmu", rect.y_emu as f64)?;
    object.set("widthEmu", rect.width_emu as f64)?;
    object.set("heightEmu", rect.height_emu as f64)?;
    Ok(object)
}

fn drawing_tree<'env>(
    env: &'env Env,
    sheet: &core::Worksheet,
) -> Result<Vec<JsObject<'env>>> {
    sheet
        .drawings()
        .iter()
        .enumerate()
        .map(|(index, object)| {
            drawing_node_to_js(
                env,
                sheet,
                &object.meta,
                DrawingPlacement::TopLevel(&object.anchor),
                &object.kind,
                &[index],
            )
        })
        .collect()
}

#[derive(Clone, Copy)]
enum DrawingFilter {
    Image,
    Chart,
    ChartEx,
    FormControl,
}

impl DrawingFilter {
    fn matches(self, kind: &core::DrawingKind) -> bool {
        matches!(
            (self, kind),
            (Self::Image, core::DrawingKind::Image(_))
                | (Self::Chart, core::DrawingKind::Chart(_))
                | (Self::ChartEx, core::DrawingKind::ChartEx(_))
                | (Self::FormControl, core::DrawingKind::FormControl(_))
        )
    }
}

fn collect_children<'env>(
    env: &'env Env,
    sheet: &core::Worksheet,
    kind: &core::DrawingKind,
    path: &[usize],
    filter: DrawingFilter,
    output: &mut Vec<JsObject<'env>>,
) -> Result<()> {
    let core::DrawingKind::Group(group) = kind else {
        return Ok(());
    };
    for (index, child) in group.children.iter().enumerate() {
        let mut child_path = path.to_vec();
        child_path.push(index);
        if filter.matches(&child.kind) {
            output.push(drawing_node_to_js(
                env,
                sheet,
                &child.meta,
                DrawingPlacement::Child(&child.transform),
                &child.kind,
                &child_path,
            )?);
        }
        collect_children(env, sheet, &child.kind, &child_path, filter, output)?;
    }
    Ok(())
}

fn filtered_drawings<'env>(
    env: &'env Env,
    sheet: &core::Worksheet,
    filter: DrawingFilter,
) -> Result<Vec<JsObject<'env>>> {
    let mut output = Vec::new();
    for (index, object) in sheet.drawings().iter().enumerate() {
        let path = [index];
        if filter.matches(&object.kind) {
            output.push(drawing_node_to_js(
                env,
                sheet,
                &object.meta,
                DrawingPlacement::TopLevel(&object.anchor),
                &object.kind,
                &path,
            )?);
        }
        collect_children(env, sheet, &object.kind, &path, filter, &mut output)?;
    }
    Ok(output)
}

fn deserialize_drawing(env: &Env, input: Unknown<'_>) -> Result<DrawingInput> {
    env.from_js_value(input)
        .map_err(|error| napi::Error::from_reason(format!("invalid drawing: {error}")))
}

/// Deserialize a `DrawingColor` DTO into a core color.
pub(crate) fn drawing_color_from_js(env: &Env, input: Unknown<'_>) -> Result<core::Color> {
    let color: DrawingColor = env
        .from_js_value(input)
        .map_err(|error| napi::Error::from_reason(format!("invalid color: {error}")))?;
    Ok(color.into())
}

fn drawing_path(path: Vec<u32>) -> Result<Vec<usize>> {
    if path.is_empty() {
        return Err(napi::Error::from_reason("drawing path cannot be empty"));
    }
    Ok(path.into_iter().map(|index| index as usize).collect())
}

fn count_drawings(sheet: &core::Worksheet, filter: DrawingFilter) -> Result<u32> {
    let count = sheet
        .drawings_flat()
        .into_iter()
        .filter(|(_, node)| filter.matches(node.kind))
        .count();
    u32::try_from(count).map_err(|_| napi::Error::from_reason("drawing count exceeds u32"))
}

fn path_starts_with(path: &[usize], prefix: &[usize]) -> bool {
    path.len() >= prefix.len() && path[..prefix.len()] == *prefix
}

/// The cell keyed by `replacement` when it is a comment. Validation
/// already rejected comments nested in groups, so only a top-level
/// comment kind can claim a cell.
// core candidate: comment-cell uniqueness belongs in core's worksheet
// drawing mutation APIs; this check is triplicated across bindings.
fn replacement_comment_cell(kind: &core::DrawingKind) -> Option<(u32, u16)> {
    match kind {
        core::DrawingKind::Comment { row, col, .. } => Some((*row, *col)),
        _ => None,
    }
}

/// Enforce one comment per cell: reject a comment `replacement` whose
/// cell already has a comment elsewhere on the sheet, ignoring
/// drawings at or under `replaced_path`.
fn ensure_comment_cells_available(
    sheet: &core::Worksheet,
    replacement: &core::DrawingKind,
    replaced_path: Option<&[usize]>,
) -> Result<()> {
    let Some(new_cell) = replacement_comment_cell(replacement) else {
        return Ok(());
    };
    for (path, node) in sheet.drawings_flat() {
        if replaced_path.is_some_and(|prefix| path_starts_with(&path, prefix)) {
            continue;
        }
        if let core::DrawingKind::Comment { row, col, .. } = node.kind {
            if (*row, *col) == new_cell {
                return Err(napi::Error::from_reason(format!(
                    "cell ({row}, {col}) already has a comment"
                )));
            }
        }
    }
    Ok(())
}

#[napi(object)]
pub struct JsFormControlInteractionResult {
    pub controls_changed: u32,
    pub linked_cells_changed: u32,
}

#[napi]
impl Worksheet {
    /// Recursive top-level drawings in back-to-front z-order.
    #[napi(getter, ts_return_type = "object[]")]
    pub fn drawings<'env>(&self, env: &'env Env) -> Result<JsObject<'env>> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            objects_to_array(env, drawing_tree(env, sheet)?)
        })
    }

    /// All form controls in depth-first drawing order.
    #[napi(getter, ts_return_type = "object[]")]
    pub fn form_controls<'env>(&self, env: &'env Env) -> Result<JsObject<'env>> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            objects_to_array(
                env,
                filtered_drawings(env, sheet, DrawingFilter::FormControl)?,
            )
        })
    }

    #[napi(getter)]
    pub fn form_control_count(&self) -> Result<u32> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            count_drawings(sheet, DrawingFilter::FormControl)
        })
    }

    /// All embedded images in depth-first drawing order, without image bytes.
    #[napi(getter, ts_return_type = "object[]")]
    pub fn images<'env>(&self, env: &'env Env) -> Result<JsObject<'env>> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            objects_to_array(
                env,
                filtered_drawings(env, sheet, DrawingFilter::Image)?,
            )
        })
    }

    #[napi(getter)]
    pub fn image_count(&self) -> Result<u32> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            count_drawings(sheet, DrawingFilter::Image)
        })
    }

    /// All standard charts in depth-first drawing order.
    #[napi(getter, ts_return_type = "object[]")]
    pub fn charts<'env>(&self, env: &'env Env) -> Result<JsObject<'env>> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            objects_to_array(
                env,
                filtered_drawings(env, sheet, DrawingFilter::Chart)?,
            )
        })
    }

    #[napi(getter)]
    pub fn chart_count(&self) -> Result<u32> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            count_drawings(sheet, DrawingFilter::Chart)
        })
    }

    /// All ChartEx charts in depth-first drawing order.
    #[napi(getter, ts_return_type = "object[]")]
    pub fn charts_ex<'env>(&self, env: &'env Env) -> Result<JsObject<'env>> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            objects_to_array(
                env,
                filtered_drawings(env, sheet, DrawingFilter::ChartEx)?,
            )
        })
    }

    #[napi(getter)]
    pub fn chart_ex_count(&self) -> Result<u32> {
        catch_panic(|| {
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            count_drawings(sheet, DrawingFilter::ChartEx)
        })
    }

    /// Append a top-level drawing and return its global z-order index.
    #[napi(ts_args_type = "input: object")]
    pub fn add_drawing(&self, env: &Env, input: Unknown<'_>) -> Result<u32> {
        catch_panic(|| {
            let object = deserialize_drawing(env, input)?
                .into_object()
                .map_err(napi::Error::from_reason)?;
            let mut workbook = self.workbook.write().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            ensure_comment_cells_available(sheet, &object.kind, None)?;
            let index = sheet.try_add_drawing(object).map_err(to_napi_err)?;
            u32::try_from(index)
                .map_err(|_| napi::Error::from_reason("drawing index exceeds u32"))
        })
    }

    /// Insert a top-level drawing at a global z-order index. Drawing
    /// paths are positional; mutating the list invalidates previously
    /// returned paths.
    #[napi(ts_args_type = "index: number, input: object")]
    pub fn insert_drawing(&self, env: &Env, index: u32, input: Unknown<'_>) -> Result<()> {
        catch_panic(|| {
            let object = deserialize_drawing(env, input)?
                .into_object()
                .map_err(napi::Error::from_reason)?;
            let mut workbook = self.workbook.write().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            ensure_comment_cells_available(sheet, &object.kind, None)?;
            sheet
                .insert_drawing(index as usize, object)
                .map_err(to_napi_err)
        })
    }

    /// Replace a top-level drawing or nested group child. Drawing
    /// paths are positional; mutating the list invalidates previously
    /// returned paths.
    #[napi(ts_args_type = "path: number[], input: object")]
    pub fn set_drawing(
        &self,
        env: &Env,
        path: Vec<u32>,
        input: Unknown<'_>,
    ) -> Result<()> {
        catch_panic(|| {
            let path = drawing_path(path)?;
            let input = deserialize_drawing(env, input)?;
            let mut workbook = self.workbook.write().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            if path.len() == 1 {
                let object = input.into_object().map_err(napi::Error::from_reason)?;
                let count = sheet.drawings().len();
                if path[0] >= count {
                    return Err(napi::Error::from_reason(format!(
                        "drawing path {path:?} out of bounds (count: {count})"
                    )));
                }
                ensure_comment_cells_available(sheet, &object.kind, Some(&path))?;
                sheet.drawings_mut()[path[0]] = object;
                return Ok(());
            }

            let child = input.into_child().map_err(napi::Error::from_reason)?;
            ensure_comment_cells_available(sheet, &child.kind, Some(&path))?;
            let (&child_index, parent_path) = path
                .split_last()
                .ok_or_else(|| napi::Error::from_reason("drawing path cannot be empty"))?;
            let parent = sheet.drawing_at_path_mut(parent_path).ok_or_else(|| {
                napi::Error::from_reason(format!("drawing path {parent_path:?} not found"))
            })?;
            let core::DrawingKind::Group(group) = parent.kind else {
                return Err(napi::Error::from_reason(
                    "nested drawing parent is not a group",
                ));
            };
            let count = group.children.len();
            let slot = group.children.get_mut(child_index).ok_or_else(|| {
                napi::Error::from_reason(format!(
                    "drawing path {path:?} out of bounds (child count: {count})"
                ))
            })?;
            *slot = child;
            Ok(())
        })
    }

    /// Remove a top-level drawing or nested group child. Drawing
    /// paths are positional; mutating the list invalidates previously
    /// returned paths.
    #[napi]
    pub fn remove_drawing(&self, path: Vec<u32>) -> Result<()> {
        catch_panic(|| {
            let path = drawing_path(path)?;
            let mut workbook = self.workbook.write().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            if path.len() == 1 {
                return sheet
                    .remove_drawing(path[0])
                    .map(|_| ())
                    .map_err(to_napi_err);
            }

            let (&child_index, parent_path) = path
                .split_last()
                .ok_or_else(|| napi::Error::from_reason("drawing path cannot be empty"))?;
            let parent = sheet.drawing_at_path_mut(parent_path).ok_or_else(|| {
                napi::Error::from_reason(format!("drawing path {parent_path:?} not found"))
            })?;
            let core::DrawingKind::Group(group) = parent.kind else {
                return Err(napi::Error::from_reason(
                    "nested drawing parent is not a group",
                ));
            };
            if child_index >= group.children.len() {
                return Err(napi::Error::from_reason(format!(
                    "drawing path {path:?} out of bounds (child count: {})",
                    group.children.len()
                )));
            }
            group.children.remove(child_index);
            Ok(())
        })
    }

    /// Move a top-level drawing to another z-order index. Drawing
    /// paths are positional; mutating the list invalidates previously
    /// returned paths.
    #[napi]
    pub fn move_drawing(&self, from: u32, to: u32) -> Result<()> {
        catch_panic(|| {
            let mut workbook = self.workbook.write().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            sheet
                .move_drawing(from as usize, to as usize)
                .map_err(to_napi_err)
        })
    }

    /// Lazily copy the bytes for an image at a drawing path. Paths are
    /// positional; mutating the drawing list invalidates previously
    /// returned paths.
    #[napi]
    pub fn drawing_image_data(&self, path: Vec<u32>) -> Result<Buffer> {
        catch_panic(|| {
            let path = drawing_path(path)?;
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            let node = sheet.drawing_at_path(&path).ok_or_else(|| {
                napi::Error::from_reason(format!("drawing path {path:?} not found"))
            })?;
            let core::DrawingKind::Image(image) = node.kind else {
                return Err(napi::Error::from_reason(
                    "drawing path does not identify an image",
                ));
            };
            Ok(image.data().to_vec().into())
        })
    }

    /// Lazily copy an image's SVG companion bytes, when present. Paths
    /// are positional; mutating the drawing list invalidates previously
    /// returned paths.
    #[napi]
    pub fn drawing_svg_data(&self, path: Vec<u32>) -> Result<Option<Buffer>> {
        catch_panic(|| {
            let path = drawing_path(path)?;
            let workbook = self.workbook.read().map_err(to_napi_err)?;
            let sheet = workbook
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;
            let node = sheet.drawing_at_path(&path).ok_or_else(|| {
                napi::Error::from_reason(format!("drawing path {path:?} not found"))
            })?;
            let core::DrawingKind::Image(image) = node.kind else {
                return Err(napi::Error::from_reason(
                    "drawing path does not identify an image",
                ));
            };
            Ok(image.svg_data().map(|bytes| bytes.to_vec().into()))
        })
    }

    /// Apply checkbox/radio semantics and synchronize linked cells immediately.
    #[napi]
    pub fn set_form_control_check_state(
        &self,
        path: Vec<u32>,
        state: JsCheckState,
    ) -> Result<JsFormControlInteractionResult> {
        catch_panic(|| {
            let path = drawing_path(path)?;
            let state = match state {
                JsCheckState::Unchecked => core::CheckState::Unchecked,
                JsCheckState::Checked => core::CheckState::Checked,
                JsCheckState::Mixed => core::CheckState::Mixed,
            };
            let mut workbook = self.workbook.write().map_err(to_napi_err)?;
            let result = workbook
                .set_form_control_check_state(self.sheet_index, &path, state)
                .map_err(to_napi_err)?;
            Ok(JsFormControlInteractionResult {
                controls_changed: u32::try_from(result.controls_changed)
                    .map_err(|_| napi::Error::from_reason("control count exceeds u32"))?,
                linked_cells_changed: u32::try_from(result.linked_cells_changed)
                    .map_err(|_| napi::Error::from_reason("linked-cell count exceeds u32"))?,
            })
        })
    }
}

#[napi]
impl Workbook {
    /// Synchronize current form-control state into linked cells.
    #[napi]
    pub fn sync_form_controls(&self) -> Result<u32> {
        catch_panic(|| {
            let mut workbook = self.inner.write().map_err(to_napi_err)?;
            u32::try_from(workbook.sync_form_control_links())
                .map_err(|_| napi::Error::from_reason("linked-cell count exceeds u32"))
        })
    }

    /// Drive form controls from formula-backed linked cells.
    #[napi]
    pub fn sync_form_controls_from_linked_cells(&self) -> Result<u32> {
        catch_panic(|| {
            let mut workbook = self.inner.write().map_err(to_napi_err)?;
            u32::try_from(workbook.sync_form_controls_from_linked_cells())
                .map_err(|_| napi::Error::from_reason("form-control count exceeds u32"))
        })
    }
}
