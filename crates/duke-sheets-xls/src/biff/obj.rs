//! Obj record (0x005D) subrecord codecs for form controls.
//!
//! An Obj record is a fixed sequence of `ft`-tagged substructures gated
//! by the object type in `FtCmo.ot` (MS-XLS 2.4.181). This module
//! provides byte-level encode/decode for the substructures that form
//! controls use:
//!
//! - `FtCmo` common object data (2.5.143)
//! - `FtCbls` checkbox/radio mirror bytes (2.5.140)
//! - `FtCblsData` checkbox/radio state (2.5.141)
//! - `FtRbo` / `FtRboData` radio grouping (2.5.152 / 2.5.153)
//! - `FtSbs` scrollable-control values (2.5.154)
//! - `FtGboData` group box properties (2.5.145)
//! - `FtLbsData` list/dropdown properties (2.5.147) with `LbsDropData`
//! - `ObjFmla` formula framing (2.5.187) wrapping an
//!   `ObjectParsedFormula` (2.5.198.22) whose rgce holds the ptg stream
//!   for cell links and input ranges
//!
//! Layout constants that the spec leaves undefined (grbit mirror bytes,
//! flag defaults) are pinned against Excel-authored BIFF8 output; the
//! unit tests carry the captured byte sequences.
//!
//! Everything here is spec-level: model conversion lives in the reader
//! and writer.

use crate::biff::parser::{read_u16, read_u8};
use crate::error::{XlsError, XlsResult};

/// Subrecord type ids (the `ft` field of each Obj substructure).
pub mod ft {
    /// Terminates an Obj record ("reserved" 4 zero bytes in MS-XLS
    /// 2.4.181; absent for list/dropdown objects).
    pub const END: u16 = 0x0000;
    /// FtMacro - macro linkage (skipped, never written).
    pub const MACRO: u16 = 0x0004;
    /// FtCf - picture clipboard format.
    pub const CF: u16 = 0x0007;
    /// FtPioGrbit - picture option flags.
    pub const PIO_GRBIT: u16 = 0x0008;
    /// FtPictFmla - picture/OLE formula.
    pub const PICT_FMLA: u16 = 0x0009;
    /// FtCbls - checkbox/radio (reserved bytes; Excel mirrors state).
    pub const CBLS: u16 = 0x000A;
    /// FtRbo - radio button (reserved bytes; Excel mirrors state).
    pub const RBO: u16 = 0x000B;
    /// FtSbs - scrollbar/spinner/list/dropdown values.
    pub const SBS: u16 = 0x000C;
    /// FtNts - note (comment) data.
    pub const NTS: u16 = 0x000D;
    /// ObjLinkFmla for scrollable controls (ot 0x10/0x11/0x12/0x14).
    pub const SBS_FMLA: u16 = 0x000E;
    /// FtGboData - group box data.
    pub const GBO_DATA: u16 = 0x000F;
    /// FtEdoData - edit box data (dialog sheets only).
    pub const EDO_DATA: u16 = 0x0010;
    /// FtRboData - radio button grouping.
    pub const RBO_DATA: u16 = 0x0011;
    /// FtCblsData - checkbox/radio state.
    pub const CBLS_DATA: u16 = 0x0012;
    /// FtLbsData - list/dropdown data.
    pub const LBS_DATA: u16 = 0x0013;
    /// ObjLinkFmla for checkbox/radio (ot 0x0B/0x0C).
    pub const CBLS_FMLA: u16 = 0x0014;
    /// FtCmo - common object data (always first).
    pub const CMO: u16 = 0x0015;
}

/// Object types (`FtCmo.ot`, MS-XLS 2.5.143).
pub mod ot {
    pub const GROUP: u16 = 0x00;
    pub const LINE: u16 = 0x01;
    pub const RECTANGLE: u16 = 0x02;
    pub const OVAL: u16 = 0x03;
    pub const ARC: u16 = 0x04;
    pub const CHART: u16 = 0x05;
    pub const TEXT: u16 = 0x06;
    pub const BUTTON: u16 = 0x07;
    pub const PICTURE: u16 = 0x08;
    pub const POLYGON: u16 = 0x09;
    pub const CHECKBOX: u16 = 0x0B;
    pub const OPTION_BUTTON: u16 = 0x0C;
    pub const EDIT_BOX: u16 = 0x0D;
    pub const LABEL: u16 = 0x0E;
    pub const DIALOG_BOX: u16 = 0x0F;
    pub const SPINNER: u16 = 0x10;
    pub const SCROLLBAR: u16 = 0x11;
    pub const LIST_BOX: u16 = 0x12;
    pub const GROUP_BOX: u16 = 0x13;
    pub const DROPDOWN: u16 = 0x14;
    pub const NOTE: u16 = 0x19;
    pub const OFFICE_ART: u16 = 0x1E;
}

/// `FtCmo.grbit` flag bits (MS-XLS 2.5.143).
pub mod cmo_flags {
    /// fLocked - object locked when the sheet is protected.
    pub const LOCKED: u16 = 0x0001;
    /// fDefaultSize.
    pub const DEFAULT_SIZE: u16 = 0x0004;
    /// fPublished.
    pub const PUBLISHED: u16 = 0x0008;
    /// fPrint - object included when printing.
    pub const PRINT: u16 = 0x0010;
    /// fDisabled.
    pub const DISABLED: u16 = 0x0080;
    /// fUIObj.
    pub const UI_OBJ: u16 = 0x0100;
    /// fRecalcObj.
    pub const RECALC_OBJ: u16 = 0x0200;
    /// fRecalcObjAlways.
    pub const RECALC_OBJ_ALWAYS: u16 = 0x1000;
    /// Undefined bit 13, set by Excel on some object kinds; mirrored
    /// for byte parity.
    pub const UNDEFINED_13: u16 = 0x2000;
    /// Undefined bit 14, set by Excel on most object kinds; mirrored
    /// for byte parity.
    pub const UNDEFINED_14: u16 = 0x4000;
}

/// FtCmo - common object data (MS-XLS 2.5.143). 22 bytes on the wire
/// including the ft/cb header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FtCmo {
    /// Object type (see [`ot`]).
    pub ot: u16,
    /// Object id, unique within the sheet substream.
    pub id: u16,
    /// Flag bits (see [`cmo_flags`]).
    pub grbit: u16,
}

impl FtCmo {
    pub fn write_to(&self, out: &mut Vec<u8>) {
        out.extend_from_slice(&ft::CMO.to_le_bytes());
        out.extend_from_slice(&0x0012u16.to_le_bytes());
        out.extend_from_slice(&self.ot.to_le_bytes());
        out.extend_from_slice(&self.id.to_le_bytes());
        out.extend_from_slice(&self.grbit.to_le_bytes());
        out.extend_from_slice(&[0u8; 12]); // unused8/9/10
    }

    /// Parse an FtCmo at `pos` (which must point at the ft field).
    /// Advances `pos` past the structure.
    pub fn read(data: &[u8], pos: &mut usize) -> XlsResult<FtCmo> {
        let ft_id = read_u16(data, pos)?;
        if ft_id != ft::CMO {
            return Err(XlsError::Parse(format!(
                "Obj record does not start with ftCmo (got {ft_id:#06x})"
            )));
        }
        let cb = read_u16(data, pos)? as usize;
        if *pos + cb > data.len() || cb < 6 {
            return Err(XlsError::Parse("ftCmo truncated".into()));
        }
        let mut p = *pos;
        let ot = read_u16(data, &mut p)?;
        let id = read_u16(data, &mut p)?;
        let grbit = read_u16(data, &mut p)?;
        *pos += cb;
        Ok(FtCmo { ot, id, grbit })
    }
}

/// Append a generic `ft`-framed subrecord.
pub fn push_subrecord(out: &mut Vec<u8>, ft_id: u16, payload: &[u8]) -> XlsResult<()> {
    let payload_len = u16::try_from(payload.len()).map_err(|_| {
        XlsError::InvalidFormat(format!(
            "Obj subrecord 0x{ft_id:04X} payload is {} bytes; maximum is {}",
            payload.len(),
            u16::MAX
        ))
    })?;
    out.extend_from_slice(&ft_id.to_le_bytes());
    out.extend_from_slice(&payload_len.to_le_bytes());
    out.extend_from_slice(payload);
    Ok(())
}

/// Append the 4-byte Obj terminator (ftEnd). Present for every object
/// type except list boxes and dropdowns (MS-XLS 2.4.181 `reserved`).
pub fn push_end(out: &mut Vec<u8>) -> XlsResult<()> {
    push_subrecord(out, ft::END, &[])
}

/// Serialize `ObjFmla` content (everything after the leading cbFmla
/// field): an `ObjectParsedFormula` (cce + 4 unused bytes + rgce)
/// padded to an even byte count. Empty rgce yields empty content
/// (cbFmla = 0).
fn obj_fmla_content(rgce: &[u8]) -> XlsResult<Vec<u8>> {
    if rgce.is_empty() {
        return Ok(Vec::new());
    }
    if rgce.len() > 0x7FFF {
        return Err(XlsError::InvalidFormat(format!(
            "ObjectParsedFormula rgce is {} bytes; maximum is 32767",
            rgce.len()
        )));
    }
    let mut content = Vec::with_capacity(8 + rgce.len());
    content.extend_from_slice(&(rgce.len() as u16).to_le_bytes());
    content.extend_from_slice(&[0u8; 4]); // ObjectParsedFormula unused
    content.extend_from_slice(rgce);
    if content.len() % 2 != 0 {
        content.push(0);
    }
    Ok(content)
}

/// Append an ObjFmla with a leading cbFmla field (the layout of
/// `FtLbsData.fmla`).
pub fn push_obj_fmla(out: &mut Vec<u8>, rgce: &[u8]) -> XlsResult<()> {
    let content = obj_fmla_content(rgce)?;
    let content_len = u16::try_from(content.len()).map_err(|_| {
        XlsError::InvalidFormat("ObjFmla content exceeds u16 length".into())
    })?;
    out.extend_from_slice(&content_len.to_le_bytes());
    out.extend_from_slice(&content);
    Ok(())
}

/// Append an `ObjLinkFmla`/`FtMacro`-style subrecord: `ft` followed by
/// an ObjFmla whose cbFmla doubles as the subrecord length field.
pub fn push_fmla_subrecord(out: &mut Vec<u8>, ft_id: u16, rgce: &[u8]) -> XlsResult<()> {
    out.extend_from_slice(&ft_id.to_le_bytes());
    push_obj_fmla(out, rgce)
}

/// Parse ObjFmla content at `pos` (pointing at cbFmla) and return the
/// rgce bytes. Advances `pos` past the whole ObjFmla.
pub fn read_obj_fmla(data: &[u8], pos: &mut usize) -> XlsResult<Vec<u8>> {
    let cb_fmla = read_u16(data, pos)? as usize;
    let Some(end) = pos.checked_add(cb_fmla) else {
        return Err(XlsError::Parse("ObjFmla length overflow".into()));
    };
    if end > data.len() {
        return Err(XlsError::Parse("ObjFmla truncated".into()));
    }
    if cb_fmla == 0 {
        return Ok(Vec::new());
    }
    let mut p = *pos;
    let cce = (read_u16(data, &mut p)? & 0x7FFF) as usize;
    p = p
        .checked_add(4)
        .ok_or_else(|| XlsError::Parse("ObjectParsedFormula offset overflow".into()))?;
    if p.checked_add(cce).is_none_or(|rgce_end| rgce_end > end) {
        *pos = end;
        return Err(XlsError::Parse("ObjectParsedFormula rgce truncated".into()));
    }
    let rgce = data[p..p + cce].to_vec();
    *pos = end;
    Ok(rgce)
}

/// Extract the rgce from ObjLinkFmla subrecord *payload* bytes (the
/// bytes after ft/cbFmla, i.e. the ObjFmla content).
fn rgce_from_fmla_payload(payload: &[u8]) -> Option<Vec<u8>> {
    if payload.len() < 6 {
        return None;
    }
    let cce = (u16::from_le_bytes([payload[0], payload[1]]) & 0x7FFF) as usize;
    payload.get(6..6 + cce).map(|s| s.to_vec())
}

/// FtSbs - scroll bar data shared by spinners, scrollbars, list boxes
/// and dropdowns (MS-XLS 2.5.154).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SbsData {
    /// Current value.
    pub val: i16,
    /// Minimum value.
    pub min: i16,
    /// Maximum value.
    pub max: i16,
    /// Minor increment (arrow click).
    pub inc: i16,
    /// Page increment (trough click).
    pub page: i16,
    /// Horizontal orientation.
    pub horizontal: bool,
    /// Scrollbar width in pixels (Excel: 22 for scrollbar/spinner,
    /// 16 for list/dropdown).
    pub dx_scroll: i16,
    /// Flag bits: 0x0001 fDraw, 0x0002 fDrawSliderOnly,
    /// 0x0004 fTrackElevator, 0x0008 fNo3d.
    pub flags: u16,
}

impl SbsData {
    pub fn write_to(&self, out: &mut Vec<u8>) -> XlsResult<()> {
        let mut payload = Vec::with_capacity(20);
        payload.extend_from_slice(&[0u8; 4]); // unused1
        payload.extend_from_slice(&self.val.to_le_bytes());
        payload.extend_from_slice(&self.min.to_le_bytes());
        payload.extend_from_slice(&self.max.to_le_bytes());
        payload.extend_from_slice(&self.inc.to_le_bytes());
        payload.extend_from_slice(&self.page.to_le_bytes());
        payload.extend_from_slice(&(self.horizontal as u16).to_le_bytes());
        payload.extend_from_slice(&self.dx_scroll.to_le_bytes());
        payload.extend_from_slice(&self.flags.to_le_bytes());
        push_subrecord(out, ft::SBS, &payload)
    }

    /// Parse from subrecord payload bytes (after ft/cb).
    pub fn from_payload(payload: &[u8]) -> XlsResult<SbsData> {
        if payload.len() < 20 {
            return Err(XlsError::Parse("ftSbs payload too short".into()));
        }
        let u = |i: usize| u16::from_le_bytes([payload[i], payload[i + 1]]);
        Ok(SbsData {
            val: u(4) as i16,
            min: u(6) as i16,
            max: u(8) as i16,
            inc: u(10) as i16,
            page: u(12) as i16,
            horizontal: u(14) != 0,
            dx_scroll: u(16) as i16,
            flags: u(18),
        })
    }
}

/// LbsDropData - dropdown-specific properties (dropdowns only).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DropData {
    /// Dropdown style: 0 combo, 1 combo edit, 2 simple dropdown.
    pub style: u16,
    /// Lines shown when dropped down.
    pub lines: u16,
    /// Minimum dropdown width in pixels (0 = automatic).
    pub min_width: u16,
}

impl Default for DropData {
    fn default() -> Self {
        DropData {
            style: 0,
            lines: 8,
            min_width: 0,
        }
    }
}

/// FtLbsData - list box / dropdown data (MS-XLS 2.5.147).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct LbsData {
    /// rgce of the input-range ObjFmla (empty = no input range).
    pub input_rgce: Vec<u8>,
    /// Number of items in the list.
    pub lines: u16,
    /// One-based index of the (first) selected item; 0 = none.
    pub sel: u16,
    /// Selection type: 0 single, 1 multi, 2 extend (wListSelType).
    pub sel_type: u16,
    /// Displayed without 3D shading.
    pub no_3d: bool,
    /// fUseCB - whether `lct` is meaningful. Excel sets this on the
    /// hidden UI dropdowns it persists for autofilter columns.
    pub use_cb: bool,
    /// Behavior class (only meaningful when `use_cb`): 0 regular
    /// dropdown, 1/7 pivot field, 3 autofilter, 5 autocomplete,
    /// 6 data validation, 9 table total row.
    pub lct: u8,
    /// Multi-selection bools, one per item (present iff sel_type != 0).
    pub multi_sel: Vec<bool>,
    /// Dropdown-specific data (present iff the object is a dropdown).
    pub drop: Option<DropData>,
}

impl LbsData {
    /// Serialize as a complete ftLbsData subrecord. `cbFContinued` is
    /// written as the actual content size, which satisfies the MS-XLS
    /// 2.5.147 requirement for structures contained in a single record.
    pub fn write_to(&self, out: &mut Vec<u8>) -> XlsResult<()> {
        if self.lines > 0x7FFF {
            return Err(XlsError::InvalidFormat(format!(
                "FtLbsData cLines {} exceeds 32767",
                self.lines
            )));
        }
        if self.sel > self.lines {
            return Err(XlsError::InvalidFormat(format!(
                "FtLbsData iSel {} exceeds cLines {}",
                self.sel, self.lines
            )));
        }
        if self.sel_type > 2 {
            return Err(XlsError::InvalidFormat(format!(
                "FtLbsData selection type {} is reserved",
                self.sel_type
            )));
        }
        if self.sel_type == 0 && !self.multi_sel.is_empty() {
            return Err(XlsError::InvalidFormat(
                "single-select FtLbsData cannot contain bsels".into(),
            ));
        }
        if self.sel_type != 0 && self.multi_sel.len() != self.lines as usize {
            return Err(XlsError::InvalidFormat(format!(
                "FtLbsData bsels has {} entries but cLines is {}",
                self.multi_sel.len(), self.lines
            )));
        }
        let mut content = Vec::new();
        push_obj_fmla(&mut content, &self.input_rgce)?;
        content.extend_from_slice(&self.lines.to_le_bytes());
        content.extend_from_slice(&self.sel.to_le_bytes());
        let flags: u16 = (self.use_cb as u16)
            | ((self.no_3d as u16) << 3)
            | ((self.sel_type & 0x3) << 4)
            | ((self.lct as u16) << 8);
        content.extend_from_slice(&flags.to_le_bytes());
        content.extend_from_slice(&0u16.to_le_bytes()); // idEdit
        if let Some(drop) = &self.drop {
            content.extend_from_slice(&drop.style.to_le_bytes());
            content.extend_from_slice(&drop.lines.to_le_bytes());
            content.extend_from_slice(&drop.min_width.to_le_bytes());
            // Empty XLUnicodeString (cch=0 + grbit=0) is 3 bytes; the
            // odd size requires one alignment byte (MS-XLS 2.5.161).
            content.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]);
        }
        if self.sel_type != 0 {
            for &b in &self.multi_sel {
                content.push(b as u8);
            }
        }
        push_subrecord(out, ft::LBS_DATA, &content)
    }

    /// Parse ftLbsData given the bytes following the ft field. The
    /// leading u16 is cbFContinued, which is unreliable as a length
    /// (Excel writes magic values), so parsing is structural. `ot`
    /// decides whether LbsDropData is present.
    pub fn parse(data: &[u8], obj_ot: u16) -> XlsResult<LbsData> {
        let mut pos = 0usize;
        let cb_continued = read_u16(data, &mut pos)?;
        if cb_continued == 0 {
            return Ok(LbsData::default());
        }
        let input_rgce = read_obj_fmla(data, &mut pos)?;
        let lines = read_u16(data, &mut pos)?;
        let sel = read_u16(data, &mut pos)?;
        let flags_lct = read_u16(data, &mut pos)?;
        let _id_edit = read_u16(data, &mut pos)?;
        let use_cb = flags_lct & 0x0001 != 0;
        let no_3d = flags_lct & 0x0008 != 0;
        let sel_type = (flags_lct >> 4) & 0x3;
        let valid_plex = flags_lct & 0x0002 != 0;
        let lct = (flags_lct >> 8) as u8;

        let mut drop = None;
        if obj_ot == ot::DROPDOWN {
            let style = read_u16(data, &mut pos)?;
            let d_lines = read_u16(data, &mut pos)?;
            let min_width = read_u16(data, &mut pos)?;
            // XLUnicodeString: cch(2) + grbit(1) + chars, plus one
            // alignment byte when the string's byte size is odd.
            let cch = read_u16(data, &mut pos)? as usize;
            let grbit = read_u8(data, &mut pos)?;
            let str_bytes = 3 + if grbit & 0x01 != 0 { cch * 2 } else { cch };
            pos += str_bytes - 3;
            if str_bytes % 2 != 0 {
                pos += 1;
            }
            drop = Some(DropData {
                style,
                lines: d_lines,
                min_width,
            });
        }
        if valid_plex {
            // rgLines: cLines XLUnicodeStrings (explicit items; dialog
            // sheet dropdowns). Skipped structurally.
            for _ in 0..lines {
                let cch = read_u16(data, &mut pos)? as usize;
                let grbit = read_u8(data, &mut pos)?;
                pos += if grbit & 0x01 != 0 { cch * 2 } else { cch };
            }
        }
        let mut multi_sel = Vec::new();
        if sel_type != 0 {
            for _ in 0..lines {
                multi_sel.push(read_u8(data, &mut pos)? != 0);
            }
        }
        Ok(LbsData {
            input_rgce,
            lines,
            sel,
            sel_type,
            no_3d,
            use_cb,
            lct,
            multi_sel,
            drop,
        })
    }
}

/// Append an FtCbls subrecord. The 12 payload bytes are reserved per
/// MS-XLS 2.5.140; Excel mirrors the checked state into the first u16
/// and writes 0x0003 in the last, so we do the same.
pub fn push_cbls(out: &mut Vec<u8>, state: u16) -> XlsResult<()> {
    let mut payload = [0u8; 12];
    payload[..2].copy_from_slice(&state.to_le_bytes());
    payload[10..].copy_from_slice(&0x0003u16.to_le_bytes());
    push_subrecord(out, ft::CBLS, &payload)
}

/// Append an FtRbo subrecord. The 6 payload bytes are reserved per
/// MS-XLS 2.5.152; Excel mirrors the checked state into the last u16.
pub fn push_rbo(out: &mut Vec<u8>, state: u16) -> XlsResult<()> {
    let mut payload = [0u8; 6];
    payload[4..].copy_from_slice(&state.to_le_bytes());
    push_subrecord(out, ft::RBO, &payload)
}

/// Append an FtCblsData subrecord (MS-XLS 2.5.141): fChecked, accel,
/// reserved, flags. Excel sets undefined flag bit 1 unconditionally;
/// mirrored for byte parity.
pub fn push_cbls_data(out: &mut Vec<u8>, state: u16, no_3d: bool) -> XlsResult<()> {
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(&state.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes()); // accel
    payload.extend_from_slice(&0u16.to_le_bytes()); // reserved
    payload.extend_from_slice(&(0x0002u16 | no_3d as u16).to_le_bytes());
    push_subrecord(out, ft::CBLS_DATA, &payload)
}

/// Append an FtRboData subrecord (MS-XLS 2.5.153).
pub fn push_rbo_data(out: &mut Vec<u8>, id_rad_next: u16, first_btn: bool) -> XlsResult<()> {
    let mut payload = Vec::with_capacity(4);
    payload.extend_from_slice(&id_rad_next.to_le_bytes());
    payload.extend_from_slice(&(first_btn as u16).to_le_bytes());
    push_subrecord(out, ft::RBO_DATA, &payload)
}

/// Append an FtGboData subrecord (MS-XLS 2.5.145).
pub fn push_gbo_data(out: &mut Vec<u8>, no_3d: bool) -> XlsResult<()> {
    let mut payload = Vec::with_capacity(6);
    payload.extend_from_slice(&0u16.to_le_bytes()); // accel
    payload.extend_from_slice(&0u16.to_le_bytes()); // reserved
    payload.extend_from_slice(&(no_3d as u16).to_le_bytes());
    push_subrecord(out, ft::GBO_DATA, &payload)
}

/// Everything extracted from one Obj record body.
#[derive(Debug, Clone, Default)]
pub struct ParsedObj {
    /// Object type from ftCmo.
    pub ot: u16,
    /// Object id from ftCmo.
    pub id: u16,
    /// Flag bits from ftCmo.
    pub grbit: u16,
    /// fChecked from FtCblsData (checkbox/radio).
    pub checked: Option<u16>,
    /// fNo3d from FtCblsData.
    pub cbls_no_3d: bool,
    /// (idRadNext, fFirstBtn) from FtRboData (radio only).
    pub radio: Option<(u16, bool)>,
    /// Scroll values from FtSbs.
    pub sbs: Option<SbsData>,
    /// rgce of the cell-link ObjLinkFmla.
    pub link_rgce: Option<Vec<u8>>,
    /// List/dropdown data from FtLbsData.
    pub lbs: Option<LbsData>,
    /// fNo3d from FtGboData (group box only).
    pub gbo_no_3d: Option<bool>,
}

/// Walk an Obj record body into a [`ParsedObj`].
///
/// Permissive: unknown subrecords are skipped by their length field;
/// a malformed subrecord terminates the walk with whatever was
/// gathered so far. `ftLbsData` must be parsed structurally (its
/// length field holds magic values in Excel output) and is always the
/// final substructure of its Obj record.
pub fn parse_obj(body: &[u8]) -> XlsResult<ParsedObj> {
    let mut pos = 0usize;
    let cmo = FtCmo::read(body, &mut pos)?;
    let mut parsed = ParsedObj {
        ot: cmo.ot,
        id: cmo.id,
        grbit: cmo.grbit,
        ..ParsedObj::default()
    };

    while pos + 4 <= body.len() {
        let ft_id = u16::from_le_bytes([body[pos], body[pos + 1]]);
        if ft_id == ft::LBS_DATA {
            if let Ok(lbs) = LbsData::parse(&body[pos + 2..], parsed.ot) {
                parsed.lbs = Some(lbs);
            }
            break;
        }
        let cb = u16::from_le_bytes([body[pos + 2], body[pos + 3]]) as usize;
        pos += 4;
        if pos + cb > body.len() {
            break;
        }
        let payload = &body[pos..pos + cb];
        match ft_id {
            ft::END => break,
            ft::CBLS_DATA => {
                if payload.len() >= 8 {
                    parsed.checked =
                        Some(u16::from_le_bytes([payload[0], payload[1]]));
                    parsed.cbls_no_3d =
                        u16::from_le_bytes([payload[6], payload[7]]) & 0x0001 != 0;
                }
            }
            ft::RBO_DATA => {
                if payload.len() >= 4 {
                    let id_rad_next = u16::from_le_bytes([payload[0], payload[1]]);
                    let first = u16::from_le_bytes([payload[2], payload[3]]) != 0;
                    parsed.radio = Some((id_rad_next, first));
                }
            }
            ft::SBS => {
                if let Ok(sbs) = SbsData::from_payload(payload) {
                    parsed.sbs = Some(sbs);
                }
            }
            ft::CBLS_FMLA | ft::SBS_FMLA => {
                parsed.link_rgce = rgce_from_fmla_payload(payload);
            }
            ft::GBO_DATA => {
                if payload.len() >= 6 {
                    parsed.gbo_no_3d =
                        Some(u16::from_le_bytes([payload[4], payload[5]]) & 0x0001 != 0);
                }
            }
            // ft::MACRO, ft::CBLS, ft::RBO, pictures, unknown: skip.
            _ => {}
        }
        pos += cb;
    }
    Ok(parsed)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Excel-authored checkbox Obj body: checked, linked to $D$2,
    /// captured from a pinned BIFF8 dump.
    const EXCEL_CHECKBOX_OBJ: &[u8] = &[
        0x15, 0x00, 0x12, 0x00, 0x0B, 0x00, 0x02, 0x00, 0x11, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // ftCmo ot=checkbox id=2 grbit=0x0011
        0x0A, 0x00, 0x0C, 0x00, 0x01, 0x00, 0x15, 0xA9, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03,
        0x00, // ftCbls with Excel garbage in the unused u16
        0x14, 0x00, 0x0C, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x24, 0x01, 0x00, 0x03, 0x00,
        0x03, // ftCblsFmla: PtgRef $D$2 (pad byte 0x03 is Excel garbage)
        0x12, 0x00, 0x08, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, // ftCblsData
        0x00, 0x00, 0x00, 0x00, // ftEnd
    ];

    /// Excel-authored option button Obj body: checked, linked to $D$3,
    /// first of its group, idRadNext=9.
    const EXCEL_OPTION_OBJ: &[u8] = &[
        0x15, 0x00, 0x12, 0x00, 0x0C, 0x00, 0x08, 0x00, 0x11, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // ftCmo ot=option id=8
        0x0A, 0x00, 0x0C, 0x00, 0x01, 0x00, 0x85, 0xA2, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03,
        0x00, // ftCbls
        0x0B, 0x00, 0x06, 0x00, 0x0D, 0xD6, 0x00, 0x00, 0x01, 0x00, // ftRbo
        0x14, 0x00, 0x0C, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x24, 0x02, 0x00, 0x03, 0x00,
        0x03, // ftCblsFmla: PtgRef $D$3
        0x12, 0x00, 0x08, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, // ftCblsData
        0x11, 0x00, 0x04, 0x00, 0x09, 0x00, 0x01, 0x00, // ftRboData idRadNext=9 fFirstBtn=1
        0x00, 0x00, 0x00, 0x00, // ftEnd
    ];

    /// Excel-authored scrollbar Obj body: val=40 min=5 max=95 inc=2
    /// page=10, linked to $D$6.
    const EXCEL_SCROLLBAR_OBJ: &[u8] = &[
        0x15, 0x00, 0x12, 0x00, 0x11, 0x00, 0x0A, 0x00, 0x11, 0x60, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // ftCmo ot=scrollbar grbit=0x6011
        0x0C, 0x00, 0x14, 0x00, 0x25, 0xA8, 0x00, 0x00, 0x28, 0x00, 0x05, 0x00, 0x5F, 0x00, 0x02,
        0x00, 0x0A, 0x00, 0x00, 0x00, 0x16, 0x00, 0x01, 0x00, // ftSbs
        0x0E, 0x00, 0x0C, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x24, 0x05, 0x00, 0x03, 0x00,
        0x03, // ftSbsFmla: PtgRef $D$6
        0x00, 0x00, 0x00, 0x00, // ftEnd
    ];

    /// Excel-authored OBJ body for the hidden dropdown Excel persists
    /// for an autofilter column: ot=0x14, grbit carries fUIObj, and
    /// FtLbsData has fUseCB=1 with lct=3 (autofilter behavior class).
    const EXCEL_AUTOFILTER_DROPDOWN_OBJ: &[u8] = &[
        0x15, 0x00, 0x12, 0x00, 0x14, 0x00, 0x01, 0x00, 0x01, 0x21, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // ftCmo ot=dropdown grbit=0x2101 (fUIObj)
        0x0C, 0x00, 0x14, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x64, 0x00, 0x01,
        0x00, 0x0A, 0x00, 0x00, 0x00, 0x10, 0x00, 0x01, 0x00, // ftSbs
        0x13, 0x00, 0xEE, 0x1F, // ftLbsData, cbFContinued magic 0x1FEE
        0x00, 0x00, // fmla: cbFmla = 0
        0x00, 0x00, // cLines
        0x04, 0x00, // iSel
        0x01, 0x03, // flags: fUseCB, lct = 0x03 (autofilter)
        0x00, 0x00, // idEdit
        0x02, 0x00, 0x08, 0x00, 0x00, 0x00, // dropData: wStyle=2 cLine=8 dxMin=0
        0x00, 0x00, 0x00, 0x00, // empty str + alignment byte
    ];

    /// Excel-authored dropdown FtLbsData bytes (after the ft field):
    /// input range $H$1:$H$4, 4 items, selection 2, 6 dropdown lines.
    const EXCEL_DROPDOWN_LBS: &[u8] = &[
        0xDE, 0x1F, // cbFContinued (Excel magic; not a length)
        0x10, 0x00, 0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x25, 0x00, 0x00, 0x03, 0x00, 0x07, 0x00,
        0x07, 0x00, 0x07, // fmla: PtgArea $H$1:$H$4 (pad 0x07 is garbage)
        0x04, 0x00, // cLines
        0x02, 0x00, // iSel
        0x08, 0x00, // flags: fNo3d
        0x00, 0x00, // idEdit
        0x00, 0x00, 0x06, 0x00, 0x00, 0x00, // dropData: wStyle=0 cLine=6 dxMin=0
        0x00, 0x00, 0x00, 0x02, // empty str + alignment byte (garbage)
    ];

    #[test]
    fn cmo_round_trips() {
        let cmo = FtCmo {
            ot: ot::CHECKBOX,
            id: 7,
            grbit: 0x0011,
        };
        let mut out = Vec::new();
        cmo.write_to(&mut out);
        assert_eq!(out.len(), 22);
        let mut pos = 0;
        let back = FtCmo::read(&out, &mut pos).unwrap();
        assert_eq!(back, cmo);
        assert_eq!(pos, 22);
    }

    #[test]
    fn obj_fmla_round_trips_and_pads() {
        // PtgRef is 5 bytes -> content 11 -> padded to 12.
        let rgce = [0x24, 0x01, 0x00, 0x03, 0x00];
        let mut out = Vec::new();
        push_obj_fmla(&mut out, &rgce).unwrap();
        assert_eq!(out.len(), 2 + 12);
        assert_eq!(u16::from_le_bytes([out[0], out[1]]), 12);
        let mut pos = 0;
        let back = read_obj_fmla(&out, &mut pos).unwrap();
        assert_eq!(back, rgce);
        assert_eq!(pos, out.len());

        // Empty rgce -> cbFmla = 0, no content.
        let mut out = Vec::new();
        push_obj_fmla(&mut out, &[]).unwrap();
        assert_eq!(out, vec![0x00, 0x00]);
        let mut pos = 0;
        assert!(read_obj_fmla(&out, &mut pos).unwrap().is_empty());
    }

    #[test]
    fn sbs_round_trips() {
        let sbs = SbsData {
            val: 40,
            min: 5,
            max: 95,
            inc: 2,
            page: 10,
            horizontal: false,
            dx_scroll: 22,
            flags: 0x0001,
        };
        let mut out = Vec::new();
        sbs.write_to(&mut out).unwrap();
        assert_eq!(out.len(), 24);
        let back = SbsData::from_payload(&out[4..]).unwrap();
        assert_eq!(back, sbs);
    }

    #[test]
    fn lbs_data_round_trips() {
        let lbs = LbsData {
            input_rgce: vec![0x25, 0x00, 0x00, 0x03, 0x00, 0x07, 0x00, 0x07, 0x00],
            lines: 4,
            sel: 2,
            sel_type: 0,
            no_3d: true,
            use_cb: false,
            lct: 0,
            multi_sel: vec![],
            drop: Some(DropData {
                style: 0,
                lines: 6,
                min_width: 0,
            }),
        };
        let mut out = Vec::new();
        lbs.write_to(&mut out).unwrap();
        let back = LbsData::parse(&out[2..], ot::DROPDOWN).unwrap();
        assert_eq!(back, lbs);
    }

    #[test]
    fn lbs_data_multi_select_round_trips() {
        let lbs = LbsData {
            input_rgce: vec![0x25, 0x00, 0x00, 0x02, 0x00, 0x07, 0x00, 0x07, 0x00],
            lines: 3,
            sel: 1,
            sel_type: 1,
            no_3d: false,
            use_cb: false,
            lct: 0,
            multi_sel: vec![true, false, true],
            drop: None,
        };
        let mut out = Vec::new();
        lbs.write_to(&mut out).unwrap();
        let back = LbsData::parse(&out[2..], ot::LIST_BOX).unwrap();
        assert_eq!(back, lbs);
    }

    #[test]
    fn parses_excel_checkbox_obj() {
        let parsed = parse_obj(EXCEL_CHECKBOX_OBJ).unwrap();
        assert_eq!(parsed.ot, ot::CHECKBOX);
        assert_eq!(parsed.id, 2);
        assert_eq!(parsed.grbit, 0x0011);
        assert_eq!(parsed.checked, Some(1));
        assert!(parsed.cbls_no_3d);
        // PtgRef row 1 col 3 = $D$2
        assert_eq!(
            parsed.link_rgce.as_deref(),
            Some(&[0x24, 0x01, 0x00, 0x03, 0x00][..])
        );
        assert!(parsed.radio.is_none());
        assert!(parsed.sbs.is_none());
    }

    #[test]
    fn parses_excel_option_obj() {
        let parsed = parse_obj(EXCEL_OPTION_OBJ).unwrap();
        assert_eq!(parsed.ot, ot::OPTION_BUTTON);
        assert_eq!(parsed.id, 8);
        assert_eq!(parsed.checked, Some(1));
        assert_eq!(parsed.radio, Some((9, true)));
        assert_eq!(
            parsed.link_rgce.as_deref(),
            Some(&[0x24, 0x02, 0x00, 0x03, 0x00][..])
        );
    }

    #[test]
    fn parses_excel_scrollbar_obj() {
        let parsed = parse_obj(EXCEL_SCROLLBAR_OBJ).unwrap();
        assert_eq!(parsed.ot, ot::SCROLLBAR);
        let sbs = parsed.sbs.expect("sbs");
        assert_eq!(sbs.val, 40);
        assert_eq!(sbs.min, 5);
        assert_eq!(sbs.max, 95);
        assert_eq!(sbs.inc, 2);
        assert_eq!(sbs.page, 10);
        assert!(!sbs.horizontal);
        assert_eq!(sbs.dx_scroll, 22);
        assert_eq!(sbs.flags, 0x0001);
        assert_eq!(
            parsed.link_rgce.as_deref(),
            Some(&[0x24, 0x05, 0x00, 0x03, 0x00][..])
        );
    }

    #[test]
    fn parses_excel_dropdown_lbs_data() {
        let lbs = LbsData::parse(EXCEL_DROPDOWN_LBS, ot::DROPDOWN).unwrap();
        assert_eq!(
            lbs.input_rgce,
            vec![0x25, 0x00, 0x00, 0x03, 0x00, 0x07, 0x00, 0x07, 0x00]
        );
        assert_eq!(lbs.lines, 4);
        assert_eq!(lbs.sel, 2);
        assert_eq!(lbs.sel_type, 0);
        assert!(lbs.no_3d);
        let drop = lbs.drop.expect("dropData");
        assert_eq!(drop.style, 0);
        assert_eq!(drop.lines, 6);
        assert_eq!(drop.min_width, 0);
    }

    #[test]
    fn parses_excel_autofilter_dropdown_obj() {
        let parsed = parse_obj(EXCEL_AUTOFILTER_DROPDOWN_OBJ).unwrap();
        assert_eq!(parsed.ot, ot::DROPDOWN);
        assert_ne!(
            parsed.grbit & super::cmo_flags::UI_OBJ,
            0,
            "autofilter dropdowns carry fUIObj"
        );
        let lbs = parsed.lbs.expect("lbs");
        assert!(lbs.use_cb);
        assert_eq!(lbs.lct, 0x03, "lct = autofilter behavior class");
        assert!(lbs.input_rgce.is_empty());
    }

    #[test]
    fn lbs_data_empty_when_cb_f_continued_zero() {
        let lbs = LbsData::parse(&[0x00, 0x00], ot::LIST_BOX).unwrap();
        assert_eq!(lbs, LbsData::default());
    }

    #[test]
    fn skips_macro_subrecord() {
        // ftCmo + ftMacro (ObjFmla with PtgRef) + ftEnd: macro must be
        // skipped via its cbFmla framing without derailing the walk.
        let mut body = Vec::new();
        FtCmo {
            ot: ot::BUTTON,
            id: 3,
            grbit: 0x4001,
        }
        .write_to(&mut body);
        push_fmla_subrecord(&mut body, ft::MACRO, &[0x23, 0x01, 0x00, 0x00, 0x00])
            .unwrap();
        push_end(&mut body).unwrap();
        let parsed = parse_obj(&body).unwrap();
        assert_eq!(parsed.ot, ot::BUTTON);
        assert!(parsed.link_rgce.is_none());
    }

    #[test]
    fn cbls_writers_mirror_excel_layout() {
        let mut out = Vec::new();
        push_cbls(&mut out, 1).unwrap();
        assert_eq!(
            out,
            vec![
                0x0A, 0x00, 0x0C, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x03, 0x00
            ]
        );

        let mut out = Vec::new();
        push_cbls_data(&mut out, 2, true).unwrap();
        assert_eq!(
            out,
            vec![0x12, 0x00, 0x08, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00]
        );

        let mut out = Vec::new();
        push_rbo(&mut out, 1).unwrap();
        assert_eq!(
            out,
            vec![0x0B, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00]
        );

        let mut out = Vec::new();
        push_rbo_data(&mut out, 9, true).unwrap();
        assert_eq!(
            out,
            vec![0x11, 0x00, 0x04, 0x00, 0x09, 0x00, 0x01, 0x00]
        );

        let mut out = Vec::new();
        push_gbo_data(&mut out, true).unwrap();
        assert_eq!(
            out,
            vec![0x0F, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00]
        );
    }

    #[test]
    fn encoder_rejects_overflowing_length_fields() {
        let mut out = Vec::new();
        let err = push_subrecord(&mut out, ft::LBS_DATA, &vec![0u8; 65_536]).unwrap_err();
        assert!(err.to_string().contains("maximum is 65535"));
        assert!(out.is_empty(), "failed framing must not mutate output");

        let err = push_obj_fmla(&mut out, &vec![0u8; 32_768]).unwrap_err();
        assert!(err.to_string().contains("maximum is 32767"));
        assert!(out.is_empty(), "failed formula framing must not mutate output");
    }

    #[test]
    fn lbs_writer_validates_selection_shape() {
        let mut out = Vec::new();
        let err = LbsData {
            lines: 3,
            sel: 0,
            sel_type: 1,
            multi_sel: vec![true, false],
            ..Default::default()
        }
        .write_to(&mut out)
        .unwrap_err();
        assert!(err.to_string().contains("2 entries but cLines is 3"));
        assert!(out.is_empty());
    }
}
