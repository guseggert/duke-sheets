//! Legacy VML drawing codec for Forms controls.
//!
//! Both XLSX and XLSB carry form controls in a legacy VML drawing
//! part (`xl/drawings/vmlDrawingN.vml`): each control is a
//! `<v:shape>` of shapetype 201 whose `<x:ClientData>` holds the
//! control state and whose `<v:textbox>` holds the caption. XLSX
//! additionally duplicates most state into `ctrlProps` parts; XLSB
//! stores controls in the VML alone.
//!
//! The emitted shape markup mirrors Excel's own output per control
//! kind (styles, fills, lock elements, ClientData element order),
//! pinned against Excel-authored files. Element semantics follow
//! ECMA-376 Part 4 (VML spreadsheet drawing extensions), with one
//! empirically-pinned quirk: `x:MoveWithCells` / `x:SizeWithCells`
//! are *negations* — the element is present when the shape does NOT
//! move / size with cells (cross-referenced against the same file's
//! `controlPr@moveWithCells/sizeWithCells` attributes).

use duke_sheets_chart::{CellMarker, DrawingAnchor, EditAs};
use duke_sheets_core::{
    CheckState, DrawingMeta, DrawingObject, FormControl, FormControlKind, ListSelection,
};
use quick_xml::events::Event;
use quick_xml::Reader;

/// EMU per screen pixel (96 dpi).
pub const EMU_PER_PX: i64 = 9525;
/// Default column width in pixels (Excel default metrics).
pub const DEFAULT_COL_PX: i64 = 64;
/// Default row height in pixels.
pub const DEFAULT_ROW_PX: i64 = 20;

/// The shapetype declaration required once per VML part containing
/// form controls (`o:spt="201"`, the host-control shape type).
pub const CONTROL_SHAPETYPE: &str = concat!(
    " <v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\"\n",
    "  path=\"m,l,21600r21600,l21600,xe\">\n",
    "  <v:stroke joinstyle=\"miter\"/>\n",
    "  <v:path shadowok=\"f\" o:extrusionok=\"f\" strokeok=\"f\" fillok=\"f\" o:connecttype=\"rect\"/>\n",
    "  <o:lock v:ext=\"edit\" shapetype=\"t\"/>\n",
    " </v:shapetype>\n",
);

/// The VML `x:ClientData/@ObjectType` name for a control kind.
pub fn vml_object_type(kind: &FormControlKind) -> &'static str {
    match kind {
        FormControlKind::Button { .. } => "Button",
        FormControlKind::Checkbox { .. } => "Checkbox",
        FormControlKind::OptionButton { .. } => "Radio",
        FormControlKind::Label { .. } => "Label",
        FormControlKind::GroupBox { .. } => "GBox",
        FormControlKind::ListBox { .. } => "List",
        FormControlKind::Dropdown { .. } => "Drop",
        FormControlKind::Scrollbar { .. } => "Scroll",
        FormControlKind::Spinner { .. } => "Spin",
    }
}

/// The 8-value `x:Anchor` tuple for a drawing anchor:
/// `[colL, dxL_px, rowT, dyT_px, colR, dxR_px, rowB, dyB_px]`.
/// One-cell and absolute anchors are extended to a cell footprint at
/// default cell metrics.
pub fn anchor_to_px(anchor: &DrawingAnchor) -> [i64; 8] {
    let (from, to) = anchor_cell_markers(anchor);
    [
        from.col as i64,
        from.col_offset_emu / EMU_PER_PX,
        from.row as i64,
        from.row_offset_emu / EMU_PER_PX,
        to.col as i64,
        to.col_offset_emu / EMU_PER_PX,
        to.row as i64,
        to.row_offset_emu / EMU_PER_PX,
    ]
}

/// Resolve any anchor variant to concrete from/to cell markers at
/// default cell metrics.
pub fn anchor_cell_markers(anchor: &DrawingAnchor) -> (CellMarker, CellMarker) {
    const COL_EMU: i128 = (DEFAULT_COL_PX * EMU_PER_PX) as i128;
    const ROW_EMU: i128 = (DEFAULT_ROW_PX * EMU_PER_PX) as i128;
    let marker_at = |x: i128, y: i128| -> CellMarker {
        let max_x = u16::MAX as i128 * COL_EMU + COL_EMU - 1;
        let max_y = u32::MAX as i128 * ROW_EMU + ROW_EMU - 1;
        let x = x.clamp(0, max_x);
        let y = y.clamp(0, max_y);
        CellMarker {
            col: (x / COL_EMU) as u16,
            col_offset_emu: (x % COL_EMU) as i64,
            row: (y / ROW_EMU) as u32,
            row_offset_emu: (y % ROW_EMU) as i64,
        }
    };
    let extend = |from: &CellMarker, width_emu: i64, height_emu: i64| -> CellMarker {
        let total_x = from.col as i128 * COL_EMU
            + from.col_offset_emu as i128
            + width_emu.max(0) as i128;
        let total_y = from.row as i128 * ROW_EMU
            + from.row_offset_emu as i128
            + height_emu.max(0) as i128;
        marker_at(total_x, total_y)
    };
    match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => (from.clone(), to.clone()),
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => (from.clone(), extend(from, *width_emu, *height_emu)),
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            let from = marker_at(*x_emu as i128, *y_emu as i128);
            let to = extend(&from, *width_emu, *height_emu);
            (from, to)
        }
    }
}

/// Rebuild a two-cell drawing anchor from an `x:Anchor` px tuple and
/// the (already un-negated) move/size flags.
pub fn px_to_anchor(a: &[i64; 8], move_with_cells: bool, size_with_cells: bool) -> DrawingAnchor {
    let edit_as = match (move_with_cells, size_with_cells) {
        (true, true) => Some(EditAs::TwoCell),
        (true, false) => Some(EditAs::OneCell),
        _ => Some(EditAs::Absolute),
    };
    DrawingAnchor::TwoCell {
        from: CellMarker {
            col: a[0].clamp(0, u16::MAX as i64) as u16,
            col_offset_emu: a[1].saturating_mul(EMU_PER_PX),
            row: a[2].clamp(0, u32::MAX as i64) as u32,
            row_offset_emu: a[3].saturating_mul(EMU_PER_PX),
        },
        to: CellMarker {
            col: a[4].clamp(0, u16::MAX as i64) as u16,
            col_offset_emu: a[5].saturating_mul(EMU_PER_PX),
            row: a[6].clamp(0, u32::MAX as i64) as u32,
            row_offset_emu: a[7].saturating_mul(EMU_PER_PX),
        },
        edit_as,
    }
}

/// The negated `x:MoveWithCells` / `x:SizeWithCells` emission for an
/// anchor's editAs semantics: `(emit_move_with_cells_element,
/// emit_size_with_cells_element)`.
fn negated_move_size(anchor: &DrawingAnchor) -> (bool, bool) {
    let edit_as = match anchor {
        DrawingAnchor::TwoCell { edit_as, .. } => edit_as.clone().unwrap_or(EditAs::TwoCell),
        DrawingAnchor::OneCell { .. } => EditAs::OneCell,
        DrawingAnchor::Absolute { .. } => EditAs::Absolute,
    };
    match edit_as {
        EditAs::TwoCell => (false, false),
        EditAs::OneCell => (false, true),
        EditAs::Absolute => (true, true),
    }
}

/// A pixel length as a pt string for VML style attributes (96 dpi ->
/// 72 dpi, trailing zeros trimmed).
pub fn px_to_pt_string(px: i64) -> String {
    let pt = px as f64 * 0.75;
    let s = format!("{pt:.2}");
    let s = s.trim_end_matches('0').trim_end_matches('.');
    s.to_string()
}

fn fmt_pt(px: i64) -> String {
    px_to_pt_string(px)
}

fn xml_escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

/// Append one control `<v:shape>` to a VML part body. `shape_id` is
/// the numeric part of `_x0000_s{id}` (must match the worksheet
/// `control/@shapeId` in XLSX); `first_button` is the recomputed
/// radio-group-head flag for option buttons. `meta` and `anchor`
/// come from the control's wrapping drawing object.
pub fn write_control_shape(
    xml: &mut String,
    shape_id: usize,
    z_index: usize,
    meta: &DrawingMeta,
    anchor: &DrawingAnchor,
    control: &FormControl,
    first_button: bool,
) {
    use FormControlKind as K;
    let kind = &control.kind;
    let a = anchor_to_px(anchor);
    let left = a[0] * DEFAULT_COL_PX + a[1];
    let top = a[2] * DEFAULT_ROW_PX + a[3];
    let width = (a[4] * DEFAULT_COL_PX + a[5]) - left;
    let height = (a[6] * DEFAULT_ROW_PX + a[7]) - top;

    let wrap_tight = matches!(
        kind,
        K::Button { .. }
            | K::Checkbox { .. }
            | K::OptionButton { .. }
            | K::GroupBox { .. }
            | K::Label { .. }
    );

    xml.push_str(&format!(
        " <v:shape id=\"_x0000_s{shape_id}\" type=\"#_x0000_t201\" style='position:absolute;\n"
    ));
    xml.push_str(&format!(
        "  margin-left:{}pt;margin-top:{}pt;width:{}pt;height:{}pt;z-index:{z_index}{}'\n",
        fmt_pt(left),
        fmt_pt(top),
        fmt_pt(width.max(0)),
        fmt_pt(height.max(0)),
        if wrap_tight { ";\n  mso-wrap-style:tight" } else { "" },
    ));
    match kind {
        K::Button { .. } => {
            xml.push_str("  fillcolor=\"buttonFace [67]\" o:insetmode=\"auto\">\n");
            xml.push_str("  <v:fill color2=\"buttonFace [67]\" o:detectmouseclick=\"t\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
        }
        K::Checkbox { .. } | K::OptionButton { .. } => {
            xml.push_str("  filled=\"f\" fillcolor=\"windowText [64]\" stroked=\"f\" strokecolor=\"window [65]\"\n");
            xml.push_str("  strokeweight=\"3e-5mm\" o:insetmode=\"auto\">\n");
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <v:path shadowok=\"t\" strokeok=\"t\" fillok=\"t\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
        }
        K::Dropdown { .. } => {
            xml.push_str(
                "  filled=\"f\" fillcolor=\"windowText [64]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n",
            );
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>\n");
        }
        K::GroupBox { .. } => {
            xml.push_str(
                "  fillcolor=\"window [65]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n",
            );
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
        }
        K::Label { .. } => {
            xml.push_str(
                "  filled=\"f\" fillcolor=\"windowText [64]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n",
            );
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
        }
        K::ListBox { .. } => {
            xml.push_str(
                "  fillcolor=\"window [65]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n",
            );
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>\n");
        }
        K::Scrollbar { .. } | K::Spinner { .. } => {
            xml.push_str("  o:insetmode=\"auto\">\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>\n");
        }
    }

    // Caption textbox for captioned kinds.
    if let Some(caption) = control.caption() {
        let (align, size) = match kind {
            K::Button { .. } => ("center", 220),
            _ => ("left", 160),
        };
        xml.push_str("  <v:textbox style='mso-direction-alt:auto' o:singleclick=\"f\">\n");
        for line in caption.split('\n') {
            let line = line.strip_suffix('\r').unwrap_or(line);
            xml.push_str(&format!(
                "   <div style='text-align:{align}'><font face=\"Segoe UI\" size=\"{size}\" color=\"auto\">{}</font></div>\n",
                xml_escape(line)
            ));
        }
        xml.push_str("  </v:textbox>\n");
    }

    // ClientData.
    xml.push_str(&format!(
        "  <x:ClientData ObjectType=\"{}\">\n",
        vml_object_type(kind)
    ));
    if !meta.locked {
        xml.push_str("   <x:Locked>False</x:Locked>\n");
    }
    let (no_move, no_size) = negated_move_size(anchor);
    if no_move {
        xml.push_str("   <x:MoveWithCells/>\n");
    }
    if no_size {
        xml.push_str("   <x:SizeWithCells/>\n");
    }
    xml.push_str(&format!(
        "   <x:Anchor>\n    {}, {}, {}, {}, {}, {}, {}, {}</x:Anchor>\n",
        a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7]
    ));
    if !meta.printable {
        xml.push_str("   <x:PrintObject>False</x:PrintObject>\n");
    }

    let push_checked = |xml: &mut String, state: &CheckState| {
        let v = match state {
            CheckState::Unchecked => 0,
            CheckState::Checked => 1,
            CheckState::Mixed => 2,
        };
        if v != 0 {
            xml.push_str(&format!("   <x:Checked>{v}</x:Checked>\n"));
        }
    };
    let push_link = |xml: &mut String, link: &Option<String>| {
        if let Some(link) = link {
            xml.push_str(&format!("   <x:FmlaLink>{}</x:FmlaLink>\n", xml_escape(link)));
        }
    };

    match kind {
        K::Button { .. } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            xml.push_str("   <x:TextHAlign>Center</x:TextHAlign>\n");
            xml.push_str("   <x:TextVAlign>Center</x:TextVAlign>\n");
        }
        K::Checkbox {
            state,
            cell_link,
            no_3d,
            ..
        } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            xml.push_str("   <x:AutoLine>False</x:AutoLine>\n");
            xml.push_str("   <x:TextVAlign>Center</x:TextVAlign>\n");
            push_checked(xml, state);
            push_link(xml, cell_link);
            if *no_3d {
                xml.push_str("   <x:NoThreeD/>\n");
            }
        }
        K::OptionButton {
            state,
            cell_link,
            no_3d,
            ..
        } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            xml.push_str("   <x:AutoLine>False</x:AutoLine>\n");
            xml.push_str("   <x:TextVAlign>Center</x:TextVAlign>\n");
            push_checked(xml, state);
            push_link(xml, cell_link);
            if *no_3d {
                xml.push_str("   <x:NoThreeD/>\n");
            }
            if first_button {
                xml.push_str("   <x:FirstButton/>\n");
            }
        }
        K::Label { .. } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
        }
        K::GroupBox { no_3d, .. } => {
            if *no_3d {
                xml.push_str("   <x:NoThreeD/>\n");
            }
        }
        K::ListBox {
            input_range,
            cell_link,
            selection,
            selected,
            no_3d,
        } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            push_link(xml, cell_link);
            xml.push_str("   <x:Val>0</x:Val>\n   <x:Min>0</x:Min>\n   <x:Max>0</x:Max>\n");
            xml.push_str("   <x:Inc>1</x:Inc>\n   <x:Page>6</x:Page>\n   <x:Dx>22</x:Dx>\n");
            if let Some(range) = input_range {
                xml.push_str(&format!(
                    "   <x:FmlaRange>{}</x:FmlaRange>\n",
                    xml_escape(range)
                ));
            }
            if let Some(&first) = selected.first() {
                xml.push_str(&format!("   <x:Sel>{}</x:Sel>\n", u32::from(first) + 1));
            }
            if *no_3d {
                xml.push_str("   <x:NoThreeD2/>\n");
            }
            let sel_type = match selection {
                ListSelection::Single => "Single",
                ListSelection::Multi => "Multi",
                ListSelection::Extend => "Extended",
            };
            xml.push_str(&format!("   <x:SelType>{sel_type}</x:SelType>\n"));
            xml.push_str("   <x:LCT>Normal</x:LCT>\n");
            if !matches!(selection, ListSelection::Single) && !selected.is_empty() {
                let list = selected
                    .iter()
                    .map(|&v| (u32::from(v) + 1).to_string())
                    .collect::<Vec<_>>()
                    .join(",");
                xml.push_str(&format!("   <x:MultiSel>{list}</x:MultiSel>\n"));
            }
        }
        K::Dropdown {
            input_range,
            cell_link,
            selected,
            lines,
            no_3d,
        } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            push_link(xml, cell_link);
            xml.push_str("   <x:Val>0</x:Val>\n   <x:Min>0</x:Min>\n   <x:Max>0</x:Max>\n");
            xml.push_str("   <x:Inc>1</x:Inc>\n   <x:Page>10</x:Page>\n   <x:Dx>22</x:Dx>\n");
            if let Some(range) = input_range {
                xml.push_str(&format!(
                    "   <x:FmlaRange>{}</x:FmlaRange>\n",
                    xml_escape(range)
                ));
            }
            if let Some(sel) = selected {
                xml.push_str(&format!("   <x:Sel>{}</x:Sel>\n", u32::from(*sel) + 1));
            }
            if *no_3d {
                xml.push_str("   <x:NoThreeD2/>\n");
            }
            xml.push_str("   <x:SelType>Single</x:SelType>\n");
            xml.push_str("   <x:LCT>Normal</x:LCT>\n");
            xml.push_str("   <x:DropStyle>Combo</x:DropStyle>\n");
            xml.push_str(&format!("   <x:DropLines>{lines}</x:DropLines>\n"));
        }
        K::Scrollbar {
            value,
            min,
            max,
            increment,
            page,
            horizontal,
            cell_link,
        } => {
            push_link(xml, cell_link);
            xml.push_str(&format!("   <x:Val>{value}</x:Val>\n"));
            xml.push_str(&format!("   <x:Min>{min}</x:Min>\n"));
            xml.push_str(&format!("   <x:Max>{max}</x:Max>\n"));
            xml.push_str(&format!("   <x:Inc>{increment}</x:Inc>\n"));
            xml.push_str(&format!("   <x:Page>{page}</x:Page>\n"));
            if *horizontal {
                xml.push_str("   <x:Horiz/>\n");
            }
            xml.push_str("   <x:Dx>22</x:Dx>\n");
        }
        K::Spinner {
            value,
            min,
            max,
            increment,
            cell_link,
        } => {
            push_link(xml, cell_link);
            xml.push_str(&format!("   <x:Val>{value}</x:Val>\n"));
            xml.push_str(&format!("   <x:Min>{min}</x:Min>\n"));
            xml.push_str(&format!("   <x:Max>{max}</x:Max>\n"));
            xml.push_str(&format!("   <x:Inc>{increment}</x:Inc>\n"));
            xml.push_str("   <x:Page>10</x:Page>\n   <x:Dx>22</x:Dx>\n");
        }
    }
    xml.push_str("  </x:ClientData>\n");
    xml.push_str(" </v:shape>\n");
}

/// One `<v:shape>`'s control-relevant contents, parsed from a VML
/// part.
#[derive(Debug, Clone, Default)]
pub struct VmlControl {
    /// Numeric shape id (the `N` of `_x0000_sN`).
    pub shape_num: u32,
    /// `x:ClientData/@ObjectType`.
    pub object_type: String,
    /// Caption text from the shape's textbox, whitespace-normalized.
    pub caption: String,
    /// `x:Anchor` values (col/px offsets).
    pub anchor_px: Option<[i64; 8]>,
    /// Un-negated flags (true = moves/sizes with cells).
    pub move_with_cells: bool,
    pub size_with_cells: bool,
    /// `x:Locked` (defaults true).
    pub locked: bool,
    /// `x:PrintObject` (defaults true).
    pub print_object: bool,
    /// `x:Checked` value (0/1/2).
    pub checked: u16,
    pub fmla_link: Option<String>,
    pub fmla_range: Option<String>,
    /// One-based `x:Sel` selection index (0 = no selection).
    pub sel: u16,
    /// `x:SelType` text (Single/Multi/Extended).
    pub sel_type: String,
    /// `x:MultiSel` one-based indices.
    pub multi_sel: Vec<u16>,
    /// `x:LCT` behavior class text (empty = unspecified).
    pub lct: String,
    pub drop_lines: u16,
    pub val: u16,
    pub min: u16,
    pub max: u16,
    pub inc: u16,
    pub page: u16,
    pub horiz: bool,
    /// `x:NoThreeD` or `x:NoThreeD2` present.
    pub no_3d: bool,
    /// `x:FirstButton` present.
    pub first_button: bool,
}

impl VmlControl {
    fn new() -> Self {
        VmlControl {
            locked: true,
            print_object: true,
            move_with_cells: true,
            size_with_cells: true,
            drop_lines: 8,
            inc: 1,
            page: 10,
            ..Default::default()
        }
    }

    /// Convert to a model [`DrawingObject`] wrapping a
    /// [`FormControl`]. Returns `None` for non-control shapes
    /// (comments) and unsupported object types.
    pub fn to_drawing_object(&self) -> Option<DrawingObject> {
        let caption = || self.caption.clone();
        let state = match self.checked {
            2 => CheckState::Mixed,
            0 => CheckState::Unchecked,
            _ => CheckState::Checked,
        };
        let kind = match self.object_type.as_str() {
            "Button" => FormControlKind::Button { caption: caption() },
            "Checkbox" => FormControlKind::Checkbox {
                caption: caption(),
                state,
                cell_link: self.fmla_link.clone(),
                no_3d: self.no_3d,
            },
            "Radio" => FormControlKind::OptionButton {
                caption: caption(),
                state: if state == CheckState::Mixed {
                    CheckState::Checked
                } else {
                    state
                },
                cell_link: self.fmla_link.clone(),
                first_in_group: self.first_button,
                no_3d: self.no_3d,
            },
            "Label" => FormControlKind::Label { caption: caption() },
            "GBox" => FormControlKind::GroupBox {
                caption: caption(),
                no_3d: self.no_3d,
            },
            "List" => {
                let selection = match self.sel_type.as_str() {
                    "Multi" => ListSelection::Multi,
                    "Extended" => ListSelection::Extend,
                    _ => ListSelection::Single,
                };
                // File values are one-based (0 = none); the model is
                // zero-based.
                let selected = if selection == ListSelection::Single {
                    if self.sel > 0 {
                        vec![self.sel - 1]
                    } else {
                        Vec::new()
                    }
                } else if !self.multi_sel.is_empty() {
                    self.multi_sel
                        .iter()
                        .filter(|&&v| v > 0)
                        .map(|&v| v - 1)
                        .collect()
                } else if self.sel > 0 {
                    vec![self.sel - 1]
                } else {
                    Vec::new()
                };
                FormControlKind::ListBox {
                    input_range: self.fmla_range.clone(),
                    cell_link: self.fmla_link.clone(),
                    selection,
                    selected,
                    no_3d: self.no_3d,
                }
            }
            "Drop" => {
                // Auxiliary UI dropdowns (autofilter, data validation)
                // carry a non-Normal behavior class.
                if !self.lct.is_empty() && self.lct != "Normal" {
                    return None;
                }
                FormControlKind::Dropdown {
                    input_range: self.fmla_range.clone(),
                    cell_link: self.fmla_link.clone(),
                    selected: if self.sel > 0 { Some(self.sel - 1) } else { None },
                    lines: self.drop_lines,
                    no_3d: self.no_3d,
                }
            }
            "Scroll" => FormControlKind::Scrollbar {
                value: self.val,
                min: self.min,
                max: self.max,
                increment: self.inc,
                page: self.page,
                horizontal: self.horiz,
                cell_link: self.fmla_link.clone(),
            },
            "Spin" => FormControlKind::Spinner {
                value: self.val,
                min: self.min,
                max: self.max,
                increment: self.inc,
                cell_link: self.fmla_link.clone(),
            },
            _ => return None,
        };

        let anchor = self
            .anchor_px
            .map(|a| px_to_anchor(&a, self.move_with_cells, self.size_with_cells))
            .unwrap_or_default();
        let mut object =
            DrawingObject::form_control(FormControl::new(kind)).with_anchor(anchor);
        object.meta.locked = self.locked;
        object.meta.printable = self.print_object;
        Some(object)
    }
}

/// One `<v:shape>` from a VML part, in document order. Legacy Note
/// (comment) shapes and control shapes share this sequence, which
/// carries their relative z-order.
#[derive(Debug, Clone)]
pub struct VmlShape {
    /// Numeric shape id (the `N` of `_x0000_sN`).
    pub shape_num: u32,
    /// Kind-specific payload.
    pub kind: VmlShapeKind,
}

/// Payload of a [`VmlShape`].
#[derive(Debug, Clone)]
pub enum VmlShapeKind {
    /// A form-control (or other non-Note) shape.
    Control(VmlControl),
    /// A cell-comment Note shape.
    Note(VmlNote),
}

/// A comment Note shape's placement data.
#[derive(Debug, Clone)]
pub struct VmlNote {
    /// Anchored cell row (`x:Row`).
    pub row: u32,
    /// Anchored cell column (`x:Column`).
    pub col: u16,
    /// `x:Anchor` values (col/px offsets).
    pub anchor_px: Option<[i64; 8]>,
    /// Popup visibility (style `visibility:visible` or `x:Visible`).
    pub visible: bool,
}

/// `x:*` boolean element semantics (ST_TrueFalseBlank): present with
/// empty/`t`/`true`/`1` text means true; `f`/`false`/`0` means false.
fn blank_true(text: &str) -> bool {
    !matches!(
        text.trim().to_ascii_lowercase().as_str(),
        "f" | "false" | "0"
    )
}

/// Collapse line-wrap whitespace: runs containing a newline become a
/// single space mid-string and vanish at the ends (they are markup
/// indentation, not caption content). Plain space runs are
/// deliberate and preserved.
fn normalize_caption(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut run = String::new();
    let flush = |out: &mut String, run: &mut String| {
        if run.is_empty() {
            return;
        }
        if run.contains('\n') || run.contains('\r') {
            if !out.is_empty() {
                out.push(' ');
            }
        } else {
            out.push_str(run);
        }
        run.clear();
    };
    for ch in text.chars() {
        if ch.is_whitespace() {
            run.push(ch);
        } else {
            flush(&mut out, &mut run);
            out.push(ch);
        }
    }
    // Trailing run: drop when it contains a newline, keep deliberate
    // trailing spaces.
    if !run.is_empty() && !run.contains('\n') && !run.contains('\r') {
        out.push_str(&run);
    }
    out
}

/// Parse every control shape out of a VML drawing part.
///
/// Permissive: parse errors (Excel embeds unclosed HTML like `<br>`
/// inside textboxes) terminate the walk but shapes completed before
/// the error are returned.
pub fn parse_vml_controls(bytes: &[u8]) -> Vec<VmlControl> {
    parse_raw_shapes(bytes)
        .into_iter()
        .map(|shape| shape.control)
        .collect()
}

/// Parse the full ordered shape sequence out of a VML drawing part:
/// control shapes and comment Note shapes, in document order.
///
/// Same permissiveness as [`parse_vml_controls`]. Note shapes lacking
/// `x:Row`/`x:Column` are dropped (they cannot be joined to a cell).
pub fn parse_vml_shapes(bytes: &[u8]) -> Vec<VmlShape> {
    parse_raw_shapes(bytes)
        .into_iter()
        .filter_map(|shape| {
            let shape_num = shape.control.shape_num;
            let kind = if shape.control.object_type == "Note" {
                let (row, col) = (shape.row?, shape.col?);
                VmlShapeKind::Note(VmlNote {
                    row,
                    col,
                    anchor_px: shape.control.anchor_px,
                    visible: shape.visible,
                })
            } else {
                VmlShapeKind::Control(shape.control)
            };
            Some(VmlShape { shape_num, kind })
        })
        .collect()
}

/// A parsed `<v:shape>` before Note/control classification.
struct RawShape {
    control: VmlControl,
    row: Option<u32>,
    col: Option<u16>,
    visible: bool,
}

fn parse_raw_shapes(bytes: &[u8]) -> Vec<RawShape> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;

    let mut out = Vec::new();
    let mut buf = Vec::new();

    let mut current: Option<RawShape> = None;
    let mut in_client_data = false;
    let mut in_textbox = false;
    let mut in_caption_div = false;
    let mut caption_line = String::new();
    let mut caption_lines: Vec<String> = Vec::new();
    let mut element_text: Option<(String, String)> = None; // (name, text)

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = e.local_name().as_ref().to_vec();
                match name.as_slice() {
                    b"shape" => {
                        let mut shape = RawShape {
                            control: VmlControl::new(),
                            row: None,
                            col: None,
                            visible: false,
                        };
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"id" => {
                                    let id = String::from_utf8_lossy(&attr.value);
                                    if let Some(num) = id.rsplit(['s', 'S']).next() {
                                        shape.control.shape_num = num.parse().unwrap_or(0);
                                    }
                                }
                                b"style" => {
                                    let style =
                                        String::from_utf8_lossy(&attr.value).to_lowercase();
                                    let normalized: String =
                                        style.chars().filter(|c| !c.is_whitespace()).collect();
                                    shape.visible = normalized.contains("visibility:visible");
                                }
                                _ => {}
                            }
                        }
                        caption_line.clear();
                        caption_lines.clear();
                        in_client_data = false;
                        in_textbox = false;
                        in_caption_div = false;
                        current = Some(shape);
                    }
                    b"textbox" => in_textbox = true,
                    b"div" if in_textbox => {
                        in_caption_div = true;
                        caption_line.clear();
                    }
                    b"ClientData" => {
                        in_client_data = true;
                        if let Some(shape) = current.as_mut() {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"ObjectType" {
                                    shape.control.object_type =
                                        String::from_utf8_lossy(&attr.value).into_owned();
                                }
                            }
                        }
                    }
                    other if in_client_data => {
                        let elem = String::from_utf8_lossy(other).into_owned();
                        element_text = Some((elem, String::new()));
                    }
                    _ => {}
                }
                // Empty elements never receive an End event; commit
                // presence-only ClientData flags immediately.
                if in_client_data {
                    if let Some(shape) = current.as_mut() {
                        match name.as_slice() {
                            b"MoveWithCells" => shape.control.move_with_cells = false,
                            b"SizeWithCells" => shape.control.size_with_cells = false,
                            b"NoThreeD" | b"NoThreeD2" => shape.control.no_3d = true,
                            b"FirstButton" => shape.control.first_button = true,
                            b"Horiz" => shape.control.horiz = true,
                            b"Visible" => shape.visible = true,
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Text(t)) => {
                let text = t
                    .unescape()
                    .map(|c| c.into_owned())
                    .unwrap_or_else(|_| String::from_utf8_lossy(t.as_ref()).into_owned());
                if in_textbox && in_caption_div {
                    caption_line.push_str(&text);
                } else if let Some((_, buf_text)) = element_text.as_mut() {
                    buf_text.push_str(&text);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.local_name().as_ref().to_vec();
                match name.as_slice() {
                    b"shape" => {
                        if let Some(mut shape) = current.take() {
                            shape.control.caption = caption_lines.join("\n");
                            if !shape.control.object_type.is_empty() {
                                out.push(shape);
                            }
                        }
                    }
                    b"textbox" => in_textbox = false,
                    b"div" => {
                        if in_caption_div {
                            caption_lines.push(normalize_caption(&caption_line));
                            caption_line.clear();
                        }
                        in_caption_div = false;
                    }
                    b"ClientData" => in_client_data = false,
                    _ => {
                        if let Some((elem, text)) = element_text.take() {
                            if elem.as_bytes() == name.as_slice() {
                                if let Some(shape) = current.as_mut() {
                                    match elem.as_str() {
                                        "Row" => shape.row = text.trim().parse().ok(),
                                        "Column" => shape.col = text.trim().parse().ok(),
                                        "Visible" => {
                                            shape.visible = blank_true(&text);
                                        }
                                        _ => apply_client_data_text(
                                            &mut shape.control,
                                            &elem,
                                            &text,
                                        ),
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break, // permissive: keep completed shapes
            _ => {}
        }
        buf.clear();
    }
    out
}

/// Apply a text-bearing ClientData element to the control being
/// built.
fn apply_client_data_text(ctrl: &mut VmlControl, elem: &str, text: &str) {
    let text = text.trim();
    let num = || text.parse::<u16>().unwrap_or(0);
    match elem {
        "Anchor" => {
            let vals: Vec<i64> = text
                .split(',')
                .filter_map(|p| p.trim().parse().ok())
                .collect();
            if vals.len() == 8 {
                let mut a = [0i64; 8];
                a.copy_from_slice(&vals);
                ctrl.anchor_px = Some(a);
            }
        }
        "Locked" => ctrl.locked = blank_true(text),
        "PrintObject" => ctrl.print_object = blank_true(text),
        "MoveWithCells" => ctrl.move_with_cells = !blank_true(text),
        "SizeWithCells" => ctrl.size_with_cells = !blank_true(text),
        "Checked" => ctrl.checked = num(),
        "FmlaLink" => ctrl.fmla_link = Some(text.to_string()),
        "FmlaRange" => ctrl.fmla_range = Some(text.to_string()),
        "Sel" => ctrl.sel = num(),
        "SelType" => ctrl.sel_type = text.to_string(),
        "LCT" => ctrl.lct = text.to_string(),
        "DropLines" => ctrl.drop_lines = num(),
        "Val" => ctrl.val = num(),
        "Min" => ctrl.min = num(),
        "Max" => ctrl.max = num(),
        "Inc" => ctrl.inc = num(),
        "Page" => ctrl.page = num(),
        "MultiSel" => {
            // Excel's emit order is not stable ("3, 1"); selection is
            // a set, so normalize to sorted indices.
            ctrl.multi_sel = text
                .split(|c: char| !c.is_ascii_digit())
                .filter(|s| !s.is_empty())
                .filter_map(|s| s.parse().ok())
                .collect();
            ctrl.multi_sel.sort_unstable();
            ctrl.multi_sel.dedup();
        }
        "Horiz" => ctrl.horiz = blank_true(text),
        "NoThreeD" | "NoThreeD2" => ctrl.no_3d = blank_true(text),
        "FirstButton" => ctrl.first_button = blank_true(text),
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use duke_sheets_core::DrawingKind;

    fn checkbox(cell_link: Option<&str>) -> DrawingObject {
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "Enable audit".to_string(),
            state: CheckState::Checked,
            cell_link: cell_link.map(str::to_string),
            no_3d: true,
        }))
        .with_anchor(DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 0,
                col_offset_emu: 13 * EMU_PER_PX,
                row: 3,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 2,
                col_offset_emu: 18 * EMU_PER_PX,
                row: 4,
                row_offset_emu: 4 * EMU_PER_PX,
            },
            edit_as: Some(EditAs::OneCell),
        })
    }

    fn shape_xml(object: &DrawingObject, shape_id: usize, z_index: usize, first: bool) -> String {
        let mut xml = String::new();
        let control = object.kind.as_form_control().expect("form control");
        write_control_shape(
            &mut xml,
            shape_id,
            z_index,
            &object.meta,
            &object.anchor,
            control,
            first,
        );
        xml
    }

    fn wrap(xml_body: &str) -> String {
        format!(
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n{xml_body}</xml>"
        )
    }

    #[test]
    fn checkbox_shape_round_trips() {
        let object = checkbox(Some("$D$2"));
        let xml = shape_xml(&object, 1026, 2, false);
        let full = wrap(&xml);

        let parsed = parse_vml_controls(full.as_bytes());
        assert_eq!(parsed.len(), 1);
        let vml = &parsed[0];
        assert_eq!(vml.shape_num, 1026);
        assert_eq!(vml.object_type, "Checkbox");
        assert_eq!(vml.caption, "Enable audit");
        assert_eq!(vml.anchor_px, Some([0, 13, 3, 0, 2, 18, 4, 4]));
        assert!(vml.move_with_cells);
        assert!(!vml.size_with_cells, "OneCell editAs negates sizing");
        assert_eq!(vml.checked, 1);
        assert_eq!(vml.fmla_link.as_deref(), Some("$D$2"));
        assert!(vml.no_3d);

        let back = vml.to_drawing_object().expect("control");
        assert_eq!(back.kind, object.kind);
        assert!(back.meta.locked);
        assert!(back.meta.printable);
    }

    #[test]
    fn excel_authored_checkbox_parses() {
        // Verbatim shape from an Excel-authored vmlDrawing part
        // (line-wrapped caption, quirky whitespace).
        let body = r##" <v:shape id="_x0000_s1026" type="#_x0000_t201" style='position:absolute;
  margin-left:9.75pt;margin-top:45pt;width:99.75pt;height:18pt;z-index:2;
  mso-wrap-style:tight' filled="f" fillcolor="windowText [64]" stroked="f"
  strokecolor="window [65]" strokeweight="3e-5mm" o:insetmode="auto">
  <v:fill color2="window [65]"/>
  <v:path shadowok="t" strokeok="t" fillok="t"/>
  <o:lock v:ext="edit" rotation="t"/>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
   <div style='text-align:left'><font face="Segoe UI" size="160" color="auto">Enable
   audit</font></div>
  </v:textbox>
  <x:ClientData ObjectType="Checkbox">
   <x:SizeWithCells/>
   <x:Anchor>
    0, 13, 3, 0, 2, 18, 4, 4</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:AutoLine>False</x:AutoLine>
   <x:TextVAlign>Center</x:TextVAlign>
   <x:Checked>1</x:Checked>
   <x:FmlaLink>$D$2</x:FmlaLink>
   <x:NoThreeD/>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        let vml = &parsed[0];
        assert_eq!(vml.caption, "Enable audit", "wrapped caption normalizes");
        assert_eq!(vml.checked, 1);
        assert!(!vml.size_with_cells);
        assert!(vml.move_with_cells);
        assert_eq!(vml.fmla_link.as_deref(), Some("$D$2"));
        let object = vml.to_drawing_object().expect("control");
        match &object.kind {
            DrawingKind::FormControl(control) => match &control.kind {
                FormControlKind::Checkbox { state, .. } => {
                    assert_eq!(state, &CheckState::Checked)
                }
                other => panic!("expected Checkbox, got {other:?}"),
            },
            other => panic!("expected form control, got {other:?}"),
        }
    }

    #[test]
    fn all_kinds_round_trip_through_vml() {
        let kinds: Vec<FormControlKind> = vec![
            FormControlKind::Button {
                caption: "Run".to_string(),
            },
            FormControlKind::OptionButton {
                caption: "Opt".to_string(),
                state: CheckState::Checked,
                cell_link: Some("$D$3".to_string()),
                first_in_group: true,
                no_3d: false,
            },
            FormControlKind::Label {
                caption: "Info <&>".to_string(),
            },
            FormControlKind::GroupBox {
                caption: "Frame".to_string(),
                no_3d: true,
            },
            FormControlKind::ListBox {
                input_range: Some("$H$1:$H$5".to_string()),
                cell_link: None,
                selection: ListSelection::Multi,
                selected: vec![0, 2, 4],
                no_3d: true,
            },
            FormControlKind::Dropdown {
                input_range: Some("$H$1:$H$4".to_string()),
                cell_link: Some("$D$4".to_string()),
                selected: Some(2),
                lines: 6,
                no_3d: true,
            },
            FormControlKind::Scrollbar {
                value: 40,
                min: 5,
                max: 95,
                increment: 2,
                page: 10,
                horizontal: true,
                cell_link: Some("$D$6".to_string()),
            },
            FormControlKind::Spinner {
                value: 12,
                min: 0,
                max: 30,
                increment: 3,
                cell_link: Some("$D$7".to_string()),
            },
        ];
        let mut xml = String::new();
        let objects: Vec<DrawingObject> = kinds
            .into_iter()
            .map(|kind| {
                DrawingObject::form_control(FormControl::new(kind)).with_anchor(
                    DrawingAnchor::TwoCell {
                        from: CellMarker {
                            col: 1,
                            col_offset_emu: 0,
                            row: 1,
                            row_offset_emu: 0,
                        },
                        to: CellMarker {
                            col: 3,
                            col_offset_emu: 0,
                            row: 3,
                            row_offset_emu: 0,
                        },
                        edit_as: None,
                    },
                )
            })
            .collect();
        for (i, object) in objects.iter().enumerate() {
            let control = object.kind.as_form_control().unwrap();
            let first = matches!(control.kind, FormControlKind::OptionButton { .. });
            xml.push_str(&shape_xml(object, 1025 + i, i + 1, first));
        }

        // Zero-based model selections serialize one-based on disk.
        assert!(xml.contains("<x:MultiSel>1,3,5</x:MultiSel>"), "{xml}");
        assert!(xml.contains("<x:Sel>3</x:Sel>"), "{xml}");

        let parsed = parse_vml_controls(wrap(&xml).as_bytes());
        assert_eq!(parsed.len(), objects.len());
        for (vml, original) in parsed.iter().zip(&objects) {
            let back = vml.to_drawing_object().expect("control");
            assert_eq!(back.kind, original.kind, "kind mismatch for {vml:?}");
        }
    }

    #[test]
    fn unlocked_unprintable_flags_round_trip() {
        let mut object = checkbox(None);
        object.meta.locked = false;
        object.meta.printable = false;
        let xml = shape_xml(&object, 1025, 1, false);
        let parsed = parse_vml_controls(wrap(&xml).as_bytes());
        let back = parsed[0].to_drawing_object().expect("control");
        assert!(!back.meta.locked);
        assert!(!back.meta.printable);
    }

    #[test]
    fn multiline_caption_round_trips_as_multiple_divs() {
        let mut object = checkbox(None);
        if let DrawingKind::FormControl(control) = &mut object.kind {
            if let FormControlKind::Checkbox { caption, .. } = &mut control.kind {
                *caption = "Line one\nLine two\n".to_string();
            }
        }
        let xml = shape_xml(&object, 1025, 1, false);
        assert_eq!(xml.matches("<div ").count(), 3);

        let parsed = parse_vml_controls(wrap(&xml).as_bytes());
        assert_eq!(parsed[0].caption, "Line one\nLine two\n");
    }

    #[test]
    fn note_shapes_are_not_controls() {
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t202">
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:Row>0</x:Row>
   <x:Column>0</x:Column>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        assert!(parsed[0].to_drawing_object().is_none());
    }

    #[test]
    fn malformed_textbox_html_keeps_completed_shapes() {
        // Unclosed <br> in a later shape must not lose earlier ones.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Checked>1</x:Checked>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <v:textbox><div>broken<br></div></v:textbox>
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 2, 0, 1, 0, 3, 0</x:Anchor>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert!(!parsed.is_empty(), "first shape must survive");
        assert_eq!(parsed[0].shape_num, 1025);
        assert_eq!(parsed[0].checked, 1);
    }

    #[test]
    fn parse_vml_shapes_keeps_note_and_control_order() {
        let body = r##" <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Checked>1</x:Checked>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1025" type="#_x0000_t202" style='position:absolute;visibility:visible'>
  <x:ClientData ObjectType="Note">
   <x:Anchor>3, 15, 1, 10, 5, 15, 5, 4</x:Anchor>
   <x:Row>2</x:Row>
   <x:Column>2</x:Column>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1027" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 4, 0, 1, 0, 5, 0</x:Anchor>
  </x:ClientData>
 </v:shape>
"##;
        let shapes = parse_vml_shapes(wrap(body).as_bytes());
        assert_eq!(shapes.len(), 3);
        assert_eq!(shapes[0].shape_num, 1026);
        assert!(matches!(shapes[0].kind, VmlShapeKind::Control(_)));
        match &shapes[1].kind {
            VmlShapeKind::Note(note) => {
                assert_eq!((note.row, note.col), (2, 2));
                assert_eq!(note.anchor_px, Some([3, 15, 1, 10, 5, 15, 5, 4]));
                assert!(note.visible);
            }
            other => panic!("expected note, got {other:?}"),
        }
        assert_eq!(shapes[2].shape_num, 1027);

        // The control-only view stays unchanged (Notes included, as
        // before, filtered downstream by to_drawing_object).
        assert_eq!(parse_vml_controls(wrap(body).as_bytes()).len(), 3);
    }

    #[test]
    fn autofilter_lct_dropdowns_are_skipped() {
        let mut vml = VmlControl::new();
        vml.object_type = "Drop".to_string();
        vml.lct = "AutoFilter".to_string();
        assert!(vml.to_drawing_object().is_none());
    }

    #[test]
    fn hostile_anchor_values_saturate_without_panicking() {
        let anchor = px_to_anchor(
            &[
                i64::MAX,
                i64::MAX,
                i64::MAX,
                i64::MIN,
                i64::MAX,
                i64::MIN,
                i64::MAX,
                i64::MAX,
            ],
            true,
            false,
        );
        match anchor {
            DrawingAnchor::TwoCell { from, to, .. } => {
                assert_eq!(from.col, u16::MAX);
                assert_eq!(from.row, u32::MAX);
                assert_eq!(from.col_offset_emu, i64::MAX);
                assert_eq!(from.row_offset_emu, i64::MIN);
                assert_eq!(to.col_offset_emu, i64::MIN);
                assert_eq!(to.row_offset_emu, i64::MAX);
            }
            other => panic!("expected TwoCell anchor, got {other:?}"),
        }
    }

    #[test]
    fn anchor_cell_markers_clamp_extreme_absolute_anchor() {
        let anchor = DrawingAnchor::Absolute {
            x_emu: i64::MAX,
            y_emu: i64::MAX,
            width_emu: i64::MAX,
            height_emu: i64::MAX,
        };
        let (from, to) = anchor_cell_markers(&anchor);
        assert_eq!(from.col, u16::MAX);
        assert_eq!(from.row, u32::MAX);
        assert_eq!(to.col, u16::MAX);
        assert_eq!(to.row, u32::MAX);
    }
}
