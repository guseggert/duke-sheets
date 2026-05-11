//! Office Art (Escher) drawing record subsystem.
//!
//! Office Art is the binary drawing layer Microsoft Office uses inside
//! BIFF8 (`.xls`), PowerPoint binary (`.ppt`), and the legacy parts of
//! Word binary (`.doc`). Inside an `.xls` file it appears as opaque
//! payload bytes inside two BIFF8 records:
//!
//! - `MSODRAWINGGROUP` (BIFF type `0x00EB`) at workbook-globals level
//!   carrying the global drawing context (default properties, shape-ID
//!   allocator state, picture blip store).
//! - `MSODRAWING` (BIFF type `0x00EC`) per sheet carrying the shapes
//!   that live on that sheet — pictures, comments (as anchored
//!   textboxes), form controls, and free-floating shapes.
//!
//! The bytes inside those BIFF records are a tree-structured format of
//! their own, called Office Art. Every node — container or leaf atom —
//! starts with the same 8-byte [`OfficeArtRecordHeader`].
//!
//! References:
//! - [MS-ODRAW] Office Drawing Binary File Format (the primary spec).
//!   Section numbers cited inline.
//! - [MS-XLS] §2.4.176 (`MsoDrawingGroup`) and §2.4.175 (`MsoDrawing`)
//!   describe the BIFF wrappers and the per-sheet shape ID rules.
//!
//! This module deliberately does not depend on anything else in
//! [`crate::biff`]. It is a self-contained encode/decode layer over
//! the Office Art byte stream. Wiring into the BIFF record stream and
//! into the `Worksheet` model lives in [`crate::writer`] and
//! [`crate::reader`].
//!
//! ## Record header layout
//!
//! [MS-ODRAW] §2.2.1 defines the 8-byte header bit packing as:
//!
//! ```text
//! offset 0  bits 0..3   recVer       (4 bits, 0xF = container)
//! offset 0  bits 4..7   recInstance, low 4 bits  (12-bit field, low nibble)
//! offset 1               recInstance, high 8 bits
//! offset 2..3            recType  (u16 LE)
//! offset 4..7            recLen   (u32 LE, length of data following the header)
//! ```
//!
//! `recLen` does **not** include the 8 header bytes themselves — it is
//! the length of the child payload.
//!
//! ## Container vs atom
//!
//! A node is a container iff `recVer == 0xF`. Containers wrap zero or
//! more child records (themselves headers + payload); the children fill
//! exactly `recLen` bytes with no padding. Atoms carry opaque payload
//! interpreted per the `recType`. See [`OfficeArtRecordHeader::is_container`].

use crate::error::{XlsError, XlsResult};

/// MS-ODRAW §2.2.1.34 OfficeArtRecType — the `recType` field of every
/// Office Art record header.
///
/// We only declare the subset this crate emits or parses. The full
/// enum has roughly 90 entries; many cover ink, animation, and chart
/// scaffolding we have no use for in BIFF8 spreadsheet drawing.
pub mod rec_type {
    /// Top-level workbook-globals container (one per `MSODRAWINGGROUP`).
    pub const DGG_CONTAINER: u16 = 0xF000;
    /// Blip (picture) store — child of `DGG_CONTAINER`.
    pub const BSTORE_CONTAINER: u16 = 0xF001;
    /// Per-drawing (per-sheet) container (one per `MSODRAWING`).
    pub const DG_CONTAINER: u16 = 0xF002;
    /// Shape-group container — wraps a group of shapes inside a `DG_CONTAINER`.
    pub const SPGR_CONTAINER: u16 = 0xF003;
    /// Single shape container — one per visible shape (picture, comment,
    /// control).
    pub const SP_CONTAINER: u16 = 0xF004;

    /// Drawing-group atom inside `DGG_CONTAINER` — global shape-ID
    /// counters and per-drawing cluster table.
    pub const FDGG: u16 = 0xF006;
    /// Blip-store-entry atom inside `BSTORE_CONTAINER` — one per stored
    /// picture.
    pub const FBSE: u16 = 0xF007;
    /// Drawing atom inside `DG_CONTAINER` — per-sheet shape count and
    /// last shape ID.
    pub const FDG: u16 = 0xF008;
    /// Shape-group atom inside `SPGR_CONTAINER` — group's logical
    /// rectangle (xLeft/yTop/xRight/yBottom).
    pub const FSPGR: u16 = 0xF009;
    /// Shape atom inside `SP_CONTAINER` — shape ID, shape type, flags.
    pub const FSP: u16 = 0xF00A;
    /// Property table atom — shape properties (fill, line, text, etc).
    pub const FOPT: u16 = 0xF00B;

    /// Client (Excel) anchor atom inside `SP_CONTAINER` — which cell
    /// the shape is anchored to.
    pub const CLIENT_ANCHOR: u16 = 0xF010;
    /// Client-data marker atom inside `SP_CONTAINER` — empty payload
    /// signalling "this shape belongs to the host application
    /// (Excel)".
    pub const CLIENT_DATA: u16 = 0xF011;
    /// Client-textbox marker atom inside `SP_CONTAINER` — empty
    /// payload signalling "this shape is a textbox; the text lives in
    /// the BIFF `TXO` record that follows".
    pub const CLIENT_TEXTBOX: u16 = 0xF00D;

    /// Default split-menu color table — child of `DGG_CONTAINER`.
    /// Always exactly 16 bytes (4 ARGB entries).
    pub const SPLIT_MENU_COLORS: u16 = 0xF11E;

    // ── Blip (image) record types (MS-ODRAW §2.2.20–2.2.25) ──────────
    /// EMF (Enhanced Metafile) blip.
    pub const BLIP_EMF: u16 = 0xF01A;
    /// WMF (Windows Metafile) blip.
    pub const BLIP_WMF: u16 = 0xF01B;
    /// PICT (Mac picture) blip.
    pub const BLIP_PICT: u16 = 0xF01C;
    /// JPEG blip.
    pub const BLIP_JPEG: u16 = 0xF01D;
    /// PNG blip.
    pub const BLIP_PNG: u16 = 0xF01E;
    /// DIB (Device-Independent Bitmap) blip.
    pub const BLIP_DIB: u16 = 0xF01F;
    /// TIFF blip.
    pub const BLIP_TIFF: u16 = 0xF029;
}

/// MS-ODRAW §2.4.24 shape type (`MSOSPT`) — the `instance` field of
/// the `FSP` atom encodes one of these values. Only the subset we
/// emit or recognise is declared.
pub mod shape_type {
    /// `msosptNotPrimitive` — used by the root `FSPGR` group shape
    /// (which is not itself a drawn primitive).
    pub const NOT_PRIMITIVE: u16 = 0x0000;
    /// `msosptRectangle` — used by some picture wrappers.
    pub const RECTANGLE: u16 = 0x0001;
    /// `msosptTextBox` — comments are textbox shapes anchored to a
    /// cell.
    pub const TEXT_BOX: u16 = 0x00CA;
    /// `msosptPictureFrame` — used by embedded pictures.
    pub const PICTURE_FRAME: u16 = 0x004B;
}

/// MS-ODRAW §2.2.32 OfficeArtFBSE blip-type codes (`btWin32`,
/// `btMacOS`, and the `rec_instance` of the embedded blip header).
pub mod blip_type {
    pub const ERROR: u8 = 0x00;
    pub const UNKNOWN: u8 = 0x01;
    pub const EMF: u8 = 0x02;
    pub const WMF: u8 = 0x03;
    pub const PICT: u8 = 0x04;
    pub const JPEG: u8 = 0x05;
    pub const PNG: u8 = 0x06;
    pub const DIB: u8 = 0x07;
    pub const TIFF: u8 = 0x11;
    pub const CMYK_JPEG: u8 = 0x12;
}

/// `rec_instance` value Excel writes for raster blip headers
/// (MS-ODRAW §2.2.23–2.2.25). The "no-secondary-uid" form has the
/// low bit clear; the "with-secondary-uid" form has it set.
pub mod blip_instance {
    /// PNG without metafileHeader / secondary UID.
    pub const PNG: u16 = 0x6E0;
    /// JPEG (no secondary UID).
    pub const JPEG: u16 = 0x46A;
    /// GIF (treated as DIB by Office binary formats).
    pub const DIB: u16 = 0x7A8;
}

/// `FSP` flag bits — MS-ODRAW §2.2.40 `OfficeArtFSP`.
pub mod fsp_flags {
    /// Bit 0: shape is a group shape (the root group, or a nested
    /// group). Mutually exclusive with most other flags.
    pub const GROUP: u32 = 0x0001;
    /// Bit 1: shape is a child of a group (i.e. lives inside an
    /// `SPGR_CONTAINER`).
    pub const CHILD: u32 = 0x0002;
    /// Bit 2: shape is the root of the per-sheet shape tree. Exactly
    /// one shape per `DG_CONTAINER` has this set, and it is the
    /// implicit group shape that sits inside the outer
    /// `SPGR_CONTAINER`.
    pub const PATRIARCH: u32 = 0x0004;
    /// Bit 3: shape has been deleted (tombstone). Excel writes these
    /// to preserve shape IDs across edits; we never emit one.
    pub const DELETED: u32 = 0x0008;
    /// Bit 4: shape has its own geometry (path), not a primitive.
    pub const OLE_SHAPE: u32 = 0x0010;
    /// Bit 5: shape has a `FOPT` table.
    pub const HAVE_MASTER: u32 = 0x0020;
    /// Bit 8: shape has a flip-horizontal applied.
    pub const FLIP_H: u32 = 0x0040;
    /// Bit 9: shape has a flip-vertical applied.
    pub const FLIP_V: u32 = 0x0080;
    /// Bit 10: shape uses a connector path (start/end attached to
    /// other shapes).
    pub const CONNECTOR: u32 = 0x0100;
    /// Bit 11: shape has an anchor record following.
    pub const HAVE_ANCHOR: u32 = 0x0200;
    /// Bit 12: shape's background fill differs from the master shape.
    pub const BACKGROUND: u32 = 0x0400;
    /// Bit 13: shape has a `FOPT` block.
    pub const HAVE_SPT: u32 = 0x0800;
}

/// MS-ODRAW §2.2.1 `OfficeArtRecordHeader` — the 8-byte header every
/// Office Art record starts with.
///
/// Use [`OfficeArtRecordHeader::write_to`] / [`OfficeArtRecordHeader::read_from`]
/// for serialisation; constructing the struct directly is also fine.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OfficeArtRecordHeader {
    /// 4-bit record version. `0xF` for containers, `0x0`..=`0xE` for atoms.
    /// MS-ODRAW typically uses `0x0` for fixed-layout atoms and `0x1`+
    /// for variable-layout atoms; we treat it as opaque on read.
    pub rec_ver: u8,
    /// 12-bit record-specific instance data. For `FSP` this is the
    /// shape type. For `SP_CONTAINER` and most containers it is 0.
    /// For `SPLIT_MENU_COLORS` it is 4. For `FOPT` it is the property
    /// count.
    pub rec_instance: u16,
    /// 16-bit record type ID. See [`rec_type`] for known values.
    pub rec_type: u16,
    /// 32-bit length of the data following the header, in bytes. Does
    /// not include the header itself.
    pub rec_len: u32,
}

/// Fixed size of every Office Art record header.
pub const HEADER_LEN: usize = 8;

impl OfficeArtRecordHeader {
    /// Construct a container header: `rec_ver = 0xF`, `rec_instance` as given.
    pub fn container(rec_type: u16, rec_instance: u16, rec_len: u32) -> Self {
        Self {
            rec_ver: 0xF,
            rec_instance,
            rec_type,
            rec_len,
        }
    }

    /// Construct an atom header.
    pub fn atom(rec_ver: u8, rec_instance: u16, rec_type: u16, rec_len: u32) -> Self {
        Self {
            rec_ver: rec_ver & 0x0F,
            rec_instance: rec_instance & 0x0FFF,
            rec_type,
            rec_len,
        }
    }

    /// MS-ODRAW §2.2.1 rule: `rec_ver == 0xF` marks a container; any
    /// other value marks an atom whose `rec_len` bytes are opaque
    /// payload.
    pub fn is_container(&self) -> bool {
        self.rec_ver == 0xF
    }

    /// Total on-disk size of this record (8-byte header + payload).
    pub fn total_len(&self) -> usize {
        HEADER_LEN + self.rec_len as usize
    }

    /// Serialise to the 8-byte on-disk header. Panics if `rec_ver`
    /// exceeds 4 bits or `rec_instance` exceeds 12 bits; use the
    /// [`OfficeArtRecordHeader::container`] /
    /// [`OfficeArtRecordHeader::atom`] constructors which clamp those
    /// fields.
    pub fn write_to(&self, out: &mut Vec<u8>) {
        debug_assert!(
            self.rec_ver <= 0x0F,
            "rec_ver overflow: {:#x}",
            self.rec_ver
        );
        debug_assert!(
            self.rec_instance <= 0x0FFF,
            "rec_instance overflow: {:#x}",
            self.rec_instance
        );
        // Byte 0: low nibble = rec_ver, high nibble = rec_instance low 4 bits.
        out.push(self.rec_ver | ((self.rec_instance & 0x000F) as u8) << 4);
        // Byte 1: rec_instance bits 11..4.
        out.push((self.rec_instance >> 4) as u8);
        // Bytes 2-3: rec_type LE.
        out.extend_from_slice(&self.rec_type.to_le_bytes());
        // Bytes 4-7: rec_len LE.
        out.extend_from_slice(&self.rec_len.to_le_bytes());
    }

    /// Parse from an 8-byte slice. Errors if the slice is too short.
    pub fn read_from(bytes: &[u8]) -> XlsResult<Self> {
        if bytes.len() < HEADER_LEN {
            return Err(XlsError::InvalidFormat(format!(
                "OfficeArt header needs {HEADER_LEN} bytes, got {}",
                bytes.len()
            )));
        }
        let rec_ver = bytes[0] & 0x0F;
        let rec_instance = (((bytes[0] >> 4) as u16) & 0x000F) | ((bytes[1] as u16) << 4);
        let rec_type = u16::from_le_bytes([bytes[2], bytes[3]]);
        let rec_len = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
        Ok(Self {
            rec_ver,
            rec_instance,
            rec_type,
            rec_len,
        })
    }
}

/// FOPT property IDs — MS-ODRAW §2.3. Only the subset the XLS writer
/// emits or the reader recognises is named here. Property IDs are
/// 14-bit identifiers; bits 14 and 15 of the on-disk `opid` are
/// `fBid` and `fComplex` respectively (see [`FoptEntry`]).
pub mod fopt_id {
    // ── Text properties (§2.3.21, ids 0x0080–0x00BF) ─────────────────
    /// `txid` (0x0080) — TXO link / text ID.
    pub const TEXT_ID: u16 = 0x0080;
    /// `dxTextLeft` (0x0081) — left inset in EMUs.
    pub const DX_TEXT_LEFT: u16 = 0x0081;
    /// `dyTextTop` (0x0082) — top inset in EMUs.
    pub const DY_TEXT_TOP: u16 = 0x0082;
    /// `dxTextRight` (0x0083) — right inset in EMUs.
    pub const DX_TEXT_RIGHT: u16 = 0x0083;
    /// `dyTextBottom` (0x0084) — bottom inset in EMUs.
    pub const DY_TEXT_BOTTOM: u16 = 0x0084;
    /// `WrapText` (0x0085) — wrap mode for the text in the shape.
    pub const WRAP_TEXT: u16 = 0x0085;
    /// `anchorText` (0x0087) — vertical anchor of the text.
    pub const ANCHOR_TEXT: u16 = 0x0087;
    /// `txflTextFlow` (0x0088) — text flow direction.
    pub const TXFL_TEXT_FLOW: u16 = 0x0088;
    /// Boolean-bag of text flags (§2.3.21.42 `gtextBooleanProperties`).
    pub const TEXT_BOOLEAN_PROPS: u16 = 0x00BF;

    // ── Fill properties (§2.3.7, ids 0x0180–0x01BF) ──────────────────
    /// `fillType` (0x0180).
    pub const FILL_TYPE: u16 = 0x0180;
    /// `fillColor` (0x0181) — primary fill colour (`OfficeArtCOLORREF`).
    pub const FILL_COLOR: u16 = 0x0181;
    /// `fillBackColor` (0x0183) — secondary fill colour.
    pub const FILL_BACK_COLOR: u16 = 0x0183;
    /// `fillCrMod` (0x0185) — fill colour mod.
    pub const FILL_CR_MOD: u16 = 0x0185;
    /// Boolean-bag of fill flags (§2.3.7.44 `FillStyleBooleanProperties`).
    pub const FILL_BOOLEAN_PROPS: u16 = 0x01BF;

    // ── Line properties (§2.3.8, ids 0x01C0–0x01FF) ──────────────────
    /// `lineColor` (0x01C0).
    pub const LINE_COLOR: u16 = 0x01C0;
    /// `lineWidth` (0x01CB) — line width in EMUs.
    pub const LINE_WIDTH: u16 = 0x01CB;
    /// Boolean-bag of line flags (§2.3.8.44 `LineStyleBooleanProperties`).
    pub const LINE_BOOLEAN_PROPS: u16 = 0x01FF;

    // ── Shadow properties (§2.3.13, ids 0x0200–0x023F) ───────────────
    /// Boolean-bag of shadow flags.
    pub const SHADOW_BOOLEAN_PROPS: u16 = 0x023F;

    // ── Shape properties (§2.3.4, ids 0x0300–0x033F) ─────────────────
    /// Boolean-bag of group-shape flags (`fHidden`, `fPrint`, etc).
    pub const GROUP_SHAPE_PROPS: u16 = 0x033F;
}

/// MS-ODRAW §2.2.40 `OfficeArtFSP` — shape descriptor atom.
///
/// One per shape, inside an `SP_CONTAINER`. Carries:
/// - The 32-bit shape ID (`spid`) — must be unique within the
///   workbook and used by the BIFF `OBJ.ftCmo.id` field to link the
///   drawing object to its host-application record.
/// - A 32-bit `grfPersistence` flag bag (see [`fsp_flags`]) describing
///   the shape's relationship to its container, deletion state, and
///   geometry source.
///
/// The shape **type** (e.g. textbox = `0x00CA`) lives in the
/// containing record header's `rec_instance` field, not in the FSP
/// body. Always 8 bytes after the header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OfficeArtFsp {
    /// 32-bit unique shape ID.
    pub spid: u32,
    /// Bitfield of [`fsp_flags`] values.
    pub grf_persistence: u32,
}

impl OfficeArtFsp {
    /// Serialise the full FSP atom (header + 8-byte body) into `out`.
    /// `shape_type` is the MSOSPT value placed in the header
    /// `rec_instance` field. Pass [`shape_type::TEXT_BOX`] for
    /// comments.
    pub fn write_to(&self, shape_type: u16, out: &mut Vec<u8>) {
        let header = OfficeArtRecordHeader::atom(2, shape_type, rec_type::FSP, 8);
        header.write_to(out);
        out.extend_from_slice(&self.spid.to_le_bytes());
        out.extend_from_slice(&self.grf_persistence.to_le_bytes());
    }

    /// Parse a full FSP atom from `bytes`. Returns `(fsp, shape_type, bytes_consumed)`.
    pub fn read_from(bytes: &[u8]) -> XlsResult<(Self, u16, usize)> {
        let header = OfficeArtRecordHeader::read_from(bytes)?;
        if header.rec_type != rec_type::FSP {
            return Err(XlsError::InvalidFormat(format!(
                "expected FSP (0x{:04X}), found 0x{:04X}",
                rec_type::FSP,
                header.rec_type
            )));
        }
        if header.rec_len != 8 {
            return Err(XlsError::InvalidFormat(format!(
                "FSP body must be 8 bytes, got {}",
                header.rec_len
            )));
        }
        if bytes.len() < HEADER_LEN + 8 {
            return Err(XlsError::InvalidFormat("FSP truncated".into()));
        }
        let spid = u32::from_le_bytes([
            bytes[HEADER_LEN],
            bytes[HEADER_LEN + 1],
            bytes[HEADER_LEN + 2],
            bytes[HEADER_LEN + 3],
        ]);
        let grf_persistence = u32::from_le_bytes([
            bytes[HEADER_LEN + 4],
            bytes[HEADER_LEN + 5],
            bytes[HEADER_LEN + 6],
            bytes[HEADER_LEN + 7],
        ]);
        Ok((
            Self {
                spid,
                grf_persistence,
            },
            header.rec_instance,
            HEADER_LEN + 8,
        ))
    }
}

/// MS-ODRAW §2.2.38 `OfficeArtFSPGR` — group-shape rectangle atom.
///
/// One per `SPGR_CONTAINER`, immediately preceding the contained
/// shapes. Defines the group's logical coordinate space as a
/// rectangle (xLeft, yTop, xRight, yBottom) in EMUs. The
/// patriarch group at the root of a sheet's drawing tree
/// conventionally uses (0, 0, 0, 0). Always 16 bytes after header.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct OfficeArtFspgr {
    pub x_left: i32,
    pub y_top: i32,
    pub x_right: i32,
    pub y_bottom: i32,
}

impl OfficeArtFspgr {
    /// Serialise the full FSPGR atom (header + 16-byte body).
    pub fn write_to(&self, out: &mut Vec<u8>) {
        let header = OfficeArtRecordHeader::atom(1, 0, rec_type::FSPGR, 16);
        header.write_to(out);
        out.extend_from_slice(&self.x_left.to_le_bytes());
        out.extend_from_slice(&self.y_top.to_le_bytes());
        out.extend_from_slice(&self.x_right.to_le_bytes());
        out.extend_from_slice(&self.y_bottom.to_le_bytes());
    }

    /// Parse a full FSPGR atom. Returns `(fspgr, bytes_consumed)`.
    pub fn read_from(bytes: &[u8]) -> XlsResult<(Self, usize)> {
        let header = OfficeArtRecordHeader::read_from(bytes)?;
        if header.rec_type != rec_type::FSPGR {
            return Err(XlsError::InvalidFormat(format!(
                "expected FSPGR (0x{:04X}), found 0x{:04X}",
                rec_type::FSPGR,
                header.rec_type
            )));
        }
        if header.rec_len != 16 || bytes.len() < HEADER_LEN + 16 {
            return Err(XlsError::InvalidFormat("FSPGR truncated".into()));
        }
        let read_i32 = |off: usize| -> i32 {
            i32::from_le_bytes([bytes[off], bytes[off + 1], bytes[off + 2], bytes[off + 3]])
        };
        Ok((
            Self {
                x_left: read_i32(HEADER_LEN),
                y_top: read_i32(HEADER_LEN + 4),
                x_right: read_i32(HEADER_LEN + 8),
                y_bottom: read_i32(HEADER_LEN + 12),
            },
            HEADER_LEN + 16,
        ))
    }
}

/// MS-XLS §2.5.193 `OfficeArtClientAnchor` for Excel — anchors a shape
/// to a cell range. 18-byte body following the 8-byte header.
///
/// Coordinate units:
/// - `dx_l` / `dx_r`: fraction of the cell's width, in 1024ths.
/// - `dy_t` / `dy_b`: fraction of the cell's height, in 256ths.
///
/// `flag` controls cell-tracking behaviour: 0 = move+size with cells,
/// 2 = move only, 3 = neither (free-floating).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct OfficeArtClientAnchor {
    /// Behaviour flags. 0 = default (move and size with cells).
    pub flag: u16,
    /// Left column.
    pub col_l: u16,
    /// X offset within left cell (1024ths of cell width).
    pub dx_l: u16,
    /// Top row.
    pub row_t: u16,
    /// Y offset within top cell (256ths of cell height).
    pub dy_t: u16,
    /// Right column (inclusive).
    pub col_r: u16,
    /// X offset within right cell.
    pub dx_r: u16,
    /// Bottom row (inclusive).
    pub row_b: u16,
    /// Y offset within bottom cell.
    pub dy_b: u16,
}

impl OfficeArtClientAnchor {
    /// Build a "comment-default" anchor that mirrors Excel's own
    /// placement for a freshly-added comment on `(row, col)`. The
    /// textbox starts one column to the right of the anchor cell on
    /// the same row and extends two columns wide by four rows tall.
    ///
    /// The exact pixel offsets (`dx_l`, `dy_t`, `dx_r`, `dy_b`) and
    /// the `flag = 3` ("do not move or size with cells") setting are
    /// what Excel itself writes — verified by driving Excel via the
    /// COM bridge and inspecting the emitted `OfficeArtClientAnchor`
    /// bytes.
    pub fn comment_default(row: u32, col: u16) -> Self {
        let row_t = row.min(u16::MAX as u32) as u16;
        let col_l = col.saturating_add(1);
        Self {
            // Excel uses flag=3 for comment anchors (msoAnchorFreeFloating).
            flag: 3,
            col_l,
            dx_l: 240,
            row_t,
            dy_t: 26,
            col_r: col_l.saturating_add(2),
            dx_r: 496,
            row_b: row_t.saturating_add(4),
            dy_b: 13,
        }
    }

    /// Serialise the full anchor atom (header + 18-byte body).
    pub fn write_to(&self, out: &mut Vec<u8>) {
        let header = OfficeArtRecordHeader::atom(0, 0, rec_type::CLIENT_ANCHOR, 18);
        header.write_to(out);
        out.extend_from_slice(&self.flag.to_le_bytes());
        out.extend_from_slice(&self.col_l.to_le_bytes());
        out.extend_from_slice(&self.dx_l.to_le_bytes());
        out.extend_from_slice(&self.row_t.to_le_bytes());
        out.extend_from_slice(&self.dy_t.to_le_bytes());
        out.extend_from_slice(&self.col_r.to_le_bytes());
        out.extend_from_slice(&self.dx_r.to_le_bytes());
        out.extend_from_slice(&self.row_b.to_le_bytes());
        out.extend_from_slice(&self.dy_b.to_le_bytes());
    }

    /// Parse a full anchor atom. Returns `(anchor, bytes_consumed)`.
    pub fn read_from(bytes: &[u8]) -> XlsResult<(Self, usize)> {
        let header = OfficeArtRecordHeader::read_from(bytes)?;
        if header.rec_type != rec_type::CLIENT_ANCHOR {
            return Err(XlsError::InvalidFormat(format!(
                "expected ClientAnchor (0x{:04X}), found 0x{:04X}",
                rec_type::CLIENT_ANCHOR,
                header.rec_type
            )));
        }
        if header.rec_len != 18 || bytes.len() < HEADER_LEN + 18 {
            return Err(XlsError::InvalidFormat("ClientAnchor truncated".into()));
        }
        let read_u16 = |off: usize| u16::from_le_bytes([bytes[off], bytes[off + 1]]);
        Ok((
            Self {
                flag: read_u16(HEADER_LEN),
                col_l: read_u16(HEADER_LEN + 2),
                dx_l: read_u16(HEADER_LEN + 4),
                row_t: read_u16(HEADER_LEN + 6),
                dy_t: read_u16(HEADER_LEN + 8),
                col_r: read_u16(HEADER_LEN + 10),
                dx_r: read_u16(HEADER_LEN + 12),
                row_b: read_u16(HEADER_LEN + 14),
                dy_b: read_u16(HEADER_LEN + 16),
            },
            HEADER_LEN + 18,
        ))
    }
}

/// MS-ODRAW §2.2.46 `OfficeArtIDCL` — one entry in an FDGG cluster
/// table. An "ID cluster" reserves a block of 1024 shape IDs for a
/// specific drawing (worksheet).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct IdCluster {
    /// Drawing ID this cluster belongs to (a 1-based per-sheet
    /// identifier matching `OfficeArtFDG.dgid`).
    pub dgid: u32,
    /// Number of shape IDs already used in this cluster.
    pub cspid_cur: u32,
}

/// MS-ODRAW §2.2.48 `OfficeArtFDGG` — drawing-group atom (one per
/// `DGG_CONTAINER`). Tracks the global shape-ID allocator state and
/// per-drawing cluster reservations.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OfficeArtFdgg {
    /// One past the highest shape ID ever used in this workbook.
    pub spid_max: u32,
    /// Total saved shapes across all drawings.
    pub csp_saved: u32,
    /// Total saved drawings (i.e. sheets that have an `MSODRAWING`).
    pub cdg_saved: u32,
    /// ID cluster reservations (one per drawing that has shapes).
    pub clusters: Vec<IdCluster>,
}

impl OfficeArtFdgg {
    /// Serialise the full FDGG atom (header + body) into `out`.
    ///
    /// MS-ODRAW writes the `cidcl` count as `clusters.len() + 1`
    /// because the array is indexed from 1; we replicate that quirk
    /// here.
    pub fn write_to(&self, out: &mut Vec<u8>) {
        // Body: 4 u32 + 8 bytes per cluster.
        let body_len = (16 + self.clusters.len() * 8) as u32;
        let header = OfficeArtRecordHeader::atom(0, 0, rec_type::FDGG, body_len);
        header.write_to(out);
        out.extend_from_slice(&self.spid_max.to_le_bytes());
        out.extend_from_slice(&((self.clusters.len() as u32) + 1).to_le_bytes());
        out.extend_from_slice(&self.csp_saved.to_le_bytes());
        out.extend_from_slice(&self.cdg_saved.to_le_bytes());
        for cluster in &self.clusters {
            out.extend_from_slice(&cluster.dgid.to_le_bytes());
            out.extend_from_slice(&cluster.cspid_cur.to_le_bytes());
        }
    }

    /// Parse a full FDGG atom from `bytes`. Returns `(fdgg, bytes_consumed)`.
    pub fn read_from(bytes: &[u8]) -> XlsResult<(Self, usize)> {
        let header = OfficeArtRecordHeader::read_from(bytes)?;
        if header.rec_type != rec_type::FDGG {
            return Err(XlsError::InvalidFormat(format!(
                "expected FDGG (0x{:04X}), found 0x{:04X}",
                rec_type::FDGG,
                header.rec_type
            )));
        }
        if header.rec_len < 16 || bytes.len() < HEADER_LEN + header.rec_len as usize {
            return Err(XlsError::InvalidFormat("FDGG truncated".into()));
        }
        let read_u32 = |off: usize| {
            u32::from_le_bytes([bytes[off], bytes[off + 1], bytes[off + 2], bytes[off + 3]])
        };
        let spid_max = read_u32(HEADER_LEN);
        let cidcl = read_u32(HEADER_LEN + 4);
        let csp_saved = read_u32(HEADER_LEN + 8);
        let cdg_saved = read_u32(HEADER_LEN + 12);

        // `cidcl` is written as `clusters.len() + 1`; reverse that.
        let cluster_count = cidcl.saturating_sub(1) as usize;
        let need = 16 + cluster_count * 8;
        if (header.rec_len as usize) < need {
            return Err(XlsError::InvalidFormat(format!(
                "FDGG declares {cluster_count} clusters but rec_len {} bytes is too short",
                header.rec_len
            )));
        }
        let mut clusters = Vec::with_capacity(cluster_count);
        for i in 0..cluster_count {
            let off = HEADER_LEN + 16 + i * 8;
            clusters.push(IdCluster {
                dgid: read_u32(off),
                cspid_cur: read_u32(off + 4),
            });
        }
        Ok((
            Self {
                spid_max,
                csp_saved,
                cdg_saved,
                clusters,
            },
            HEADER_LEN + header.rec_len as usize,
        ))
    }
}

/// MS-ODRAW §2.2.49 `OfficeArtFDG` — per-drawing atom (one per
/// `DG_CONTAINER`). Tracks the shape count and last shape ID used in
/// this drawing (sheet).
///
/// The `dgid` (drawing ID) is encoded in the containing header's
/// `rec_instance`, not in the FDG body.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OfficeArtFdg {
    /// Number of shapes saved in this drawing (including the
    /// patriarch).
    pub csp_saved: u32,
    /// One past the highest shape ID used in this drawing.
    pub spid_last: u32,
}

impl OfficeArtFdg {
    /// Serialise the full FDG atom (header + 8 body bytes) into
    /// `out`. `dgid` is placed in the header `rec_instance` per
    /// MS-ODRAW §2.2.49.
    pub fn write_to(&self, dgid: u16, out: &mut Vec<u8>) {
        let header = OfficeArtRecordHeader::atom(0, dgid, rec_type::FDG, 8);
        header.write_to(out);
        out.extend_from_slice(&self.csp_saved.to_le_bytes());
        out.extend_from_slice(&self.spid_last.to_le_bytes());
    }

    /// Parse a full FDG atom. Returns `(fdg, dgid, bytes_consumed)`.
    pub fn read_from(bytes: &[u8]) -> XlsResult<(Self, u16, usize)> {
        let header = OfficeArtRecordHeader::read_from(bytes)?;
        if header.rec_type != rec_type::FDG {
            return Err(XlsError::InvalidFormat(format!(
                "expected FDG (0x{:04X}), found 0x{:04X}",
                rec_type::FDG,
                header.rec_type
            )));
        }
        if header.rec_len != 8 || bytes.len() < HEADER_LEN + 8 {
            return Err(XlsError::InvalidFormat("FDG truncated".into()));
        }
        let csp_saved = u32::from_le_bytes([
            bytes[HEADER_LEN],
            bytes[HEADER_LEN + 1],
            bytes[HEADER_LEN + 2],
            bytes[HEADER_LEN + 3],
        ]);
        let spid_last = u32::from_le_bytes([
            bytes[HEADER_LEN + 4],
            bytes[HEADER_LEN + 5],
            bytes[HEADER_LEN + 6],
            bytes[HEADER_LEN + 7],
        ]);
        Ok((
            Self {
                csp_saved,
                spid_last,
            },
            header.rec_instance,
            HEADER_LEN + 8,
        ))
    }
}

/// MS-ODRAW §2.2.83 `OfficeArtSplitMenuColorContainer` — the default
/// 16-byte palette (4 ARGB entries) that appears in every
/// `DGG_CONTAINER`. Excel's defaults are the standard colour-picker
/// fill, line, shadow, and 3-D colours.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SplitMenuColors {
    pub fill_color: u32,
    pub line_color: u32,
    pub shadow_color: u32,
    pub three_d_color: u32,
}

impl SplitMenuColors {
    /// The exact 16-byte palette Excel emits for new workbooks.
    /// Values verified by driving Excel via the COM bridge and
    /// inspecting the `OfficeArtSplitMenuColorContainer` bytes inside
    /// `MSODRAWINGGROUP`.
    ///
    /// All four `u32` slots are `OfficeArtCOLORREF` values. The high
    /// byte selects the colour type (`0x08` = scheme colour,
    /// `0x10` = system colour); the low 24 bits index into the
    /// referenced palette.
    pub const EXCEL_DEFAULT: Self = Self {
        fill_color: 0x0800_000D,
        line_color: 0x0800_000C,
        shadow_color: 0x0800_0017,
        three_d_color: 0x1000_00F7,
    };

    /// Serialise the full atom (header + 16-byte body).
    pub fn write_to(&self, out: &mut Vec<u8>) {
        // Per MS-ODRAW the rec_instance is 4 (entry count).
        let header = OfficeArtRecordHeader::atom(0, 4, rec_type::SPLIT_MENU_COLORS, 16);
        header.write_to(out);
        out.extend_from_slice(&self.fill_color.to_le_bytes());
        out.extend_from_slice(&self.line_color.to_le_bytes());
        out.extend_from_slice(&self.shadow_color.to_le_bytes());
        out.extend_from_slice(&self.three_d_color.to_le_bytes());
    }
}

/// Write a generic Office Art container record: an 8-byte header
/// (`rec_ver=0xF`) wrapping the given `body` bytes. Used by callers
/// that have already serialised the container's contents and need
/// only the header.
///
/// This single helper covers `SP_CONTAINER`, `SPGR_CONTAINER`,
/// `DG_CONTAINER`, `DGG_CONTAINER`, and `BSTORE_CONTAINER`. The
/// `rec_instance` field is record-specific:
/// - `DG_CONTAINER` uses the drawing id.
/// - `BSTORE_CONTAINER` uses the blip-store entry count.
/// - Other containers use 0.
pub fn write_container(rec_type_id: u16, rec_instance: u16, body: &[u8], out: &mut Vec<u8>) {
    let header = OfficeArtRecordHeader::container(rec_type_id, rec_instance, body.len() as u32);
    header.write_to(out);
    out.extend_from_slice(body);
}

/// Build an `SP_CONTAINER` for a comment textbox shape, anchored to
/// `(row, col)`, with the FOPT properties given. Emits the full
/// container (header + FSP + FOPT + ClientAnchor + ClientData +
/// ClientTextbox) into `out`.
///
/// `spid` is the unique shape ID; the caller is responsible for
/// allocating it via a [`ShapeIdAllocator`]-style counter.
pub fn write_comment_sp_container(
    spid: u32,
    row: u32,
    col: u16,
    properties: &FoptTable,
    out: &mut Vec<u8>,
) {
    let mut body = Vec::new();
    // FSP: textbox, with anchor + master flags.
    OfficeArtFsp {
        spid,
        grf_persistence: fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT,
    }
    .write_to(shape_type::TEXT_BOX, &mut body);
    // FOPT (caller-supplied).
    properties.write_to(&mut body);
    // ClientAnchor at the requested cell.
    OfficeArtClientAnchor::comment_default(row, col).write_to(&mut body);
    // ClientData marker.
    write_client_data(&mut body);
    // ClientTextbox marker — text content is in the BIFF TXO that
    // follows the MSODRAWING record.
    write_client_textbox(&mut body);

    write_container(rec_type::SP_CONTAINER, 0, &body, out);
}

/// MS-ODRAW §2.2.23 `OfficeArtBlipPNG` (and analogous variants for
/// JPEG / DIB / etc.). Carries one image's compressed file bytes
/// preceded by a 16-byte MD4/MD5 UID and a 1-byte tag.
///
/// Layout after the 8-byte record header (recVer=0,
/// recInstance=`PNG`/`JPEG`/..., recType=`BLIP_PNG`/...):
///
/// ```text
/// rgbUid     : [u8; 16]   // matches the FBSE's rgbUid
/// tag        : u8         // 0xFF
/// blip_data  : [u8]       // raw image file bytes
/// ```
#[derive(Debug, Clone)]
pub struct OfficeArtBlip {
    /// Record type — one of the `BLIP_*` constants in [`rec_type`].
    pub rec_type: u16,
    /// `rec_instance` value placed in the header — one of the
    /// `blip_instance` constants.
    pub rec_instance: u16,
    /// 16-byte unique identifier of the image content. Conventionally
    /// the MD4 hash of `data`, but Excel does not validate the value,
    /// so any deterministic 16-byte sequence works.
    pub rgb_uid: [u8; 16],
    /// Image file bytes (PNG, JPEG, etc.).
    pub data: Vec<u8>,
}

impl OfficeArtBlip {
    /// Construct a PNG blip with the MD5-derived UID we use for
    /// rgbUid. Excel itself uses an MD4 hash but does not validate
    /// the value, so any deterministic 16 bytes are accepted.
    pub fn png(data: Vec<u8>) -> Self {
        Self {
            rec_type: rec_type::BLIP_PNG,
            rec_instance: blip_instance::PNG,
            rgb_uid: rgb_uid_for(&data),
            data,
        }
    }

    /// Construct a JPEG blip. Same UID scheme as PNG.
    pub fn jpeg(data: Vec<u8>) -> Self {
        Self {
            rec_type: rec_type::BLIP_JPEG,
            rec_instance: blip_instance::JPEG,
            rgb_uid: rgb_uid_for(&data),
            data,
        }
    }

    /// Construct a DIB (Device-Independent Bitmap) blip. Used for
    /// BMP input. The caller must pass the BMP **without** the
    /// 14-byte `BITMAPFILEHEADER` — the DIB blip stores only the
    /// `BITMAPINFOHEADER` + optional palette + pixel data.
    pub fn dib(data: Vec<u8>) -> Self {
        Self {
            rec_type: rec_type::BLIP_DIB,
            rec_instance: blip_instance::DIB,
            rgb_uid: rgb_uid_for(&data),
            data,
        }
    }

    /// Total size of this blip on disk (header + body).
    pub fn total_size(&self) -> u32 {
        HEADER_LEN as u32 + 16 + 1 + self.data.len() as u32
    }

    /// Body size (excluding the 8-byte record header).
    pub fn body_size(&self) -> u32 {
        16 + 1 + self.data.len() as u32
    }

    /// Serialise the full blip (header + UID + tag + image data) into `out`.
    pub fn write_to(&self, out: &mut Vec<u8>) {
        let header =
            OfficeArtRecordHeader::atom(0, self.rec_instance, self.rec_type, self.body_size());
        header.write_to(out);
        out.extend_from_slice(&self.rgb_uid);
        out.push(0xFF);
        out.extend_from_slice(&self.data);
    }
}

/// MS-ODRAW §2.2.32 `OfficeArtFBSE` — one entry in the `BSTORE_CONTAINER`
/// blip-store. Wraps an [`OfficeArtBlip`] plus per-entry metadata
/// (compression type, reference count, optional delayed-loading offset).
///
/// `rec_ver=2`, `rec_instance` = `bt_win32` blip-type code (the same
/// value as the `bt_win32` field — Excel duplicates this in the
/// header instance for fast lookup).
#[derive(Debug, Clone)]
pub struct OfficeArtFbse {
    /// MS-ODRAW blip-type code (one of [`blip_type`] values) — also
    /// placed in the header `rec_instance`.
    pub bt_win32: u8,
    /// Mac OS blip-type code. We mirror `bt_win32` since we don't
    /// emit different Mac variants.
    pub bt_mac_os: u8,
    /// Embedded blip carrying the image bytes. `foDelay=0` so the
    /// blip immediately follows the FBSE header.
    pub blip: OfficeArtBlip,
    /// Reference count. Defaults to 1 (one shape uses this blip).
    pub c_ref: u32,
}

impl OfficeArtFbse {
    /// Construct an FBSE wrapping the given blip with reference
    /// count 1 and Win32/MacOS blip-type both set to the blip's
    /// canonical code (PNG=6, JPEG=5, etc.).
    pub fn new(blip: OfficeArtBlip) -> Self {
        let bt = blip_type_for_rec_type(blip.rec_type);
        Self {
            bt_win32: bt,
            bt_mac_os: bt,
            blip,
            c_ref: 1,
        }
    }

    /// Total on-disk size (header + 36 fixed body bytes + blip total).
    pub fn total_size(&self) -> u32 {
        HEADER_LEN as u32 + 36 + self.blip.total_size()
    }

    /// Serialise the full FBSE record into `out`.
    pub fn write_to(&self, out: &mut Vec<u8>) {
        let body_len = 36 + self.blip.total_size();
        let header = OfficeArtRecordHeader::atom(2, self.bt_win32 as u16, rec_type::FBSE, body_len);
        header.write_to(out);
        out.push(self.bt_win32);
        out.push(self.bt_mac_os);
        out.extend_from_slice(&self.blip.rgb_uid);
        out.extend_from_slice(&[0xFF, 0x00]); // tag (constant)
        out.extend_from_slice(&self.blip.total_size().to_le_bytes());
        out.extend_from_slice(&self.c_ref.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // foDelay (immediate)
        out.push(0); // usage
        out.push(0); // cbName (no name)
        out.push(0); // unused2
        out.push(0); // unused3
                     // No nameData (cbName == 0). Embedded blip follows.
        self.blip.write_to(out);
    }
}

/// Map an `OfficeArtBlip` record type to its `OfficeArtFBSE` blip
/// type code (the value that goes in `btWin32` / `btMacOS` and the
/// FBSE header's `rec_instance`).
fn blip_type_for_rec_type(rt: u16) -> u8 {
    match rt {
        rec_type::BLIP_EMF => blip_type::EMF,
        rec_type::BLIP_WMF => blip_type::WMF,
        rec_type::BLIP_PICT => blip_type::PICT,
        rec_type::BLIP_JPEG => blip_type::JPEG,
        rec_type::BLIP_PNG => blip_type::PNG,
        rec_type::BLIP_DIB => blip_type::DIB,
        rec_type::BLIP_TIFF => blip_type::TIFF,
        _ => blip_type::UNKNOWN,
    }
}

/// Compute a 16-byte unique identifier for an image. MS-ODRAW
/// specifies MD4 here, but Excel does not validate the value — any
/// deterministic 16-byte hash works. We use the first 16 bytes of
/// the image's MD5, which is plenty for de-duplication purposes
/// inside a single workbook.
pub fn rgb_uid_for(data: &[u8]) -> [u8; 16] {
    use md5::{Digest, Md5};
    let hash = Md5::digest(data);
    let mut out = [0u8; 16];
    out.copy_from_slice(&hash[..16]);
    out
}

/// Build a `BSTORE_CONTAINER` carrying the given FBSE entries.
/// Emits the full container (8-byte header + serialised FBSEs)
/// into `out`. The header's `rec_instance` is set to the entry
/// count, matching what Excel writes.
pub fn write_bstore_container(fbses: &[OfficeArtFbse], out: &mut Vec<u8>) {
    let mut body = Vec::new();
    for fbse in fbses {
        fbse.write_to(&mut body);
    }
    let header = OfficeArtRecordHeader::container(
        rec_type::BSTORE_CONTAINER,
        fbses.len() as u16,
        body.len() as u32,
    );
    header.write_to(out);
    out.extend_from_slice(&body);
}

/// Build the FOPT property table Excel writes for a picture-frame
/// shape, modelled on the empirical 8-entry pattern observed via
/// the COM bridge probe.
///
/// `blip_id` is the 1-based index into the workbook's blip store —
/// the `pib` (picture blip index) FOPT property points at this.
/// `shape_name` is the user-visible name (e.g. "Picture 1") stored
/// in the `wzName` complex property.
pub fn picture_fopt(blip_id: u32, shape_name: &str) -> FoptTable {
    picture_fopt_with(blip_id, shape_name, None)
}

/// Picture FOPT with an optional rotation. `rotation` is in 60,000ths
/// of a degree (the OOXML / OfficeArt unit). `None` omits the
/// `0x0004` rotation property entirely (matching Excel's emit for
/// pictures with no rotation).
///
/// FOPT entries must appear in ascending `id` order; the rotation
/// property is inserted before the existing `0x007F` protection
/// entry when present.
pub fn picture_fopt_with(blip_id: u32, shape_name: &str, rotation: Option<i32>) -> FoptTable {
    let mut t = FoptTable::new();

    // 0x0004: rotation (60,000ths of a degree). Goes first because
    // FOPT requires ascending opid order. We only emit this entry
    // if rotation is set; absence means "no rotation".
    if let Some(rot) = rotation {
        t.push(FoptEntry::simple(0x0004, rot as u32));
    }

    // 0x007F: protection booleans (lockAspectRatio etc.).
    t.push(FoptEntry::simple(0x007F, 0x01FB_0000));
    // 0x0104: pib (picture blip index, fBid flag set).
    t.push(FoptEntry::blip_id(0x0104, blip_id));
    // 0x013F: geometry booleans.
    t.push(FoptEntry::simple(0x013F, 0x0006_0000));
    // 0x01BF: fill booleans.
    t.push(FoptEntry::simple(0x01BF, 0x0010_0000));
    // 0x01FF: line booleans.
    t.push(FoptEntry::simple(0x01FF, 0x0008_0000));
    // 0x033F: shape booleans (fPrint, fHidden, …).
    t.push(FoptEntry::simple(0x033F, 0x0018_0010));

    // 0x0380: wzName (shape name) — complex property with UTF-16LE
    // payload + trailing null. Excel sets BOTH the fBid and
    // fComplex bits on this opid; we mirror that exactly so the
    // bytes match.
    let mut name_bytes: Vec<u8> = shape_name
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect();
    name_bytes.extend_from_slice(&[0, 0]); // null terminator
    t.push(FoptEntry {
        id: 0x0380,
        is_blip_id: true, // Excel sets fBid on the wzName entry
        value: FoptValue::Complex(name_bytes),
    });

    // 0x03BF: group/shape booleans.
    t.push(FoptEntry::simple(0x03BF, 0x0002_0000));
    t
}

/// Build an `SP_CONTAINER` for a picture-frame shape. Emits the full
/// container (header + FSP + FOPT + ClientAnchor + ClientData)
/// into `out`. Pictures do NOT include a ClientTextbox marker —
/// they have no associated TXO record.
pub fn write_picture_sp_container(
    spid: u32,
    blip_id: u32,
    shape_name: &str,
    anchor: OfficeArtClientAnchor,
    out: &mut Vec<u8>,
) {
    let mut body = Vec::new();
    OfficeArtFsp {
        spid,
        grf_persistence: fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT,
    }
    .write_to(shape_type::PICTURE_FRAME, &mut body);
    picture_fopt(blip_id, shape_name).write_to(&mut body);
    anchor.write_to(&mut body);
    write_client_data(&mut body);
    write_container(rec_type::SP_CONTAINER, 0, &body, out);
}

/// Build the patriarch `SP_CONTAINER` for a sheet's drawing — the
/// implicit root group shape. Exactly one per `SPGR_CONTAINER` at
/// the top of the sheet's drawing tree.
///
/// Body = FSPGR (zero rectangle) + FSP (GROUP|PATRIARCH flags,
/// shape_type = NOT_PRIMITIVE).
pub fn write_patriarch_sp_container(spid: u32, out: &mut Vec<u8>) {
    let mut body = Vec::new();
    // FSPGR with zero rectangle.
    OfficeArtFspgr::default().write_to(&mut body);
    // FSP marking this shape as the per-sheet patriarch group.
    OfficeArtFsp {
        spid,
        grf_persistence: fsp_flags::GROUP | fsp_flags::PATRIARCH,
    }
    .write_to(shape_type::NOT_PRIMITIVE, &mut body);

    write_container(rec_type::SP_CONTAINER, 0, &body, out);
}

/// MS-ODRAW §2.2.31 `OfficeArtClientData` — empty marker atom that
/// signals "the following BIFF records (`OBJ`, `TXO`, etc.) describe
/// the host-application data for this shape". Body is zero bytes.
pub fn write_client_data(out: &mut Vec<u8>) {
    let header = OfficeArtRecordHeader::atom(0, 0, rec_type::CLIENT_DATA, 0);
    header.write_to(out);
}

/// MS-ODRAW §2.2.34 `OfficeArtClientTextbox` — empty marker atom that
/// signals "this shape has text; the text content lives in the next
/// BIFF `TXO` record". Body is zero bytes.
pub fn write_client_textbox(out: &mut Vec<u8>) {
    let header = OfficeArtRecordHeader::atom(0, 0, rec_type::CLIENT_TEXTBOX, 0);
    header.write_to(out);
}

/// One entry in an Office Art FOPT property table (MS-ODRAW §2.3.1
/// `OfficeArtFOPTE`).
///
/// On disk every entry is exactly 6 bytes:
///
/// ```text
///   bits 0..13   property id
///   bit 14       fBid     (op is a blip ID rather than a literal)
///   bit 15       fComplex (op is the length of a trailing complex payload)
///   bits 16..47  op       (value, blip id, or complex-payload length)
/// ```
///
/// We model this as a separate [`FoptEntry::value`] enum so callers do
/// not have to manage the `fComplex` bit by hand. The on-disk packing
/// is performed by [`FoptTable::write_to`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FoptEntry {
    /// 14-bit property id (see [`fopt_id`] for known names).
    pub id: u16,
    /// `fBid` flag — `true` when `value` is a blip ID rather than a
    /// literal value. Only meaningful for simple entries.
    pub is_blip_id: bool,
    /// Property value. Simple entries pack a `u32` into the `op`
    /// field; complex entries store opaque bytes following all
    /// entries.
    pub value: FoptValue,
}

/// Body of an [`FoptEntry`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FoptValue {
    /// 4-byte `op` field is the literal value.
    Simple(u32),
    /// `op` is the length of these bytes; the bytes themselves are
    /// stored in the complex-data trailer after the entry array.
    Complex(Vec<u8>),
}

impl FoptEntry {
    /// Construct a simple-valued entry (`fComplex=0`, `fBid=0`).
    pub fn simple(id: u16, value: u32) -> Self {
        Self {
            id: id & 0x3FFF,
            is_blip_id: false,
            value: FoptValue::Simple(value),
        }
    }

    /// Construct a blip-id-valued entry (`fComplex=0`, `fBid=1`).
    pub fn blip_id(id: u16, value: u32) -> Self {
        Self {
            id: id & 0x3FFF,
            is_blip_id: true,
            value: FoptValue::Simple(value),
        }
    }

    /// Construct a complex-valued entry (`fComplex=1`).
    pub fn complex(id: u16, value: Vec<u8>) -> Self {
        Self {
            id: id & 0x3FFF,
            is_blip_id: false,
            value: FoptValue::Complex(value),
        }
    }
}

/// Build the 14-entry `FoptTable` Excel writes for a freshly-created
/// cell comment. The exact property IDs and values are what Excel
/// itself emits — verified by driving Excel via the COM bridge and
/// inspecting the `FOPT` payload on a comment textbox shape.
///
/// `text_id` is the per-shape `txid` placed in the `TEXT_ID` slot
/// (property `0x0080`). Excel writes a different value for every
/// shape; we use a per-comment counter so the values stay distinct
/// across a workbook.
pub fn comment_fopt(text_id: u32) -> FoptTable {
    let mut t = FoptTable::new();
    t.push(FoptEntry::simple(fopt_id::TEXT_ID, text_id));
    t.push(FoptEntry::simple(0x008B, 0x0000_0002)); // related text props
    t.push(FoptEntry::simple(fopt_id::TEXT_BOOLEAN_PROPS, 0x0008_0008));
    t.push(FoptEntry::simple(0x0158, 0)); // reserved text prop
    t.push(FoptEntry::simple(fopt_id::FILL_COLOR, 0x0800_0050));
    t.push(FoptEntry::simple(fopt_id::FILL_BACK_COLOR, 0x0800_0050));
    t.push(FoptEntry::simple(fopt_id::FILL_CR_MOD, 0x1000_00F4));
    t.push(FoptEntry::simple(fopt_id::FILL_BOOLEAN_PROPS, 0x0010_0010));
    t.push(FoptEntry::simple(fopt_id::LINE_COLOR, 0x0800_0051));
    t.push(FoptEntry::simple(0x01C3, 0x1000_00F4)); // line colour mod
    t.push(FoptEntry::simple(0x0201, 0x0800_0051)); // shadow colour
    t.push(FoptEntry::simple(0x0203, 0x1000_00F4)); // shadow colour mod
    t.push(FoptEntry::simple(
        fopt_id::SHADOW_BOOLEAN_PROPS,
        0x0003_0001,
    ));
    t.push(FoptEntry::simple(fopt_id::GROUP_SHAPE_PROPS, 0x0002_0002));
    t
}

/// Build the 3-entry `FoptTable` Excel writes inside the workbook's
/// `MSODRAWINGGROUP` (the global default property table). Values
/// verified empirically.
pub fn dgg_default_fopt() -> FoptTable {
    let mut t = FoptTable::new();
    t.push(FoptEntry::simple(fopt_id::TEXT_BOOLEAN_PROPS, 0x0008_0008));
    t.push(FoptEntry::simple(fopt_id::FILL_COLOR, 0x0800_0041));
    t.push(FoptEntry::simple(fopt_id::LINE_COLOR, 0x0800_0040));
    t
}

/// Office Art property table (MS-ODRAW §2.3.1 `OfficeArtFOPT`).
///
/// Wraps an ordered list of [`FoptEntry`] values plus an
/// implicit trailing complex-data block. The [`Self::write_to`] /
/// [`Self::read_from`] helpers serialise to and from the full atom
/// including its 8-byte [`OfficeArtRecordHeader`].
///
/// Property order matters: MS-ODRAW requires entries be sorted in
/// ascending `id` order. Callers should add properties in numeric
/// order; [`Self::sort_entries`] enforces this before serialisation.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct FoptTable {
    entries: Vec<FoptEntry>,
}

impl FoptTable {
    /// Construct an empty table.
    pub fn new() -> Self {
        Self::default()
    }

    /// Number of entries in the table.
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// `true` iff the table has no entries.
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Append an entry.
    pub fn push(&mut self, entry: FoptEntry) {
        self.entries.push(entry);
    }

    /// Iterate over entries in their current order.
    pub fn entries(&self) -> impl Iterator<Item = &FoptEntry> {
        self.entries.iter()
    }

    /// Sort entries by ascending `id`, as required by MS-ODRAW for
    /// the on-disk form. Callers that add entries in ascending id
    /// order can skip this; [`Self::write_to`] does NOT auto-sort,
    /// so misordered tables emit misordered bytes.
    pub fn sort_entries(&mut self) {
        self.entries.sort_by_key(|e| e.id);
    }

    /// Total payload length (entry array + complex-data trailer),
    /// not including the 8-byte record header.
    fn payload_len(&self) -> u32 {
        let entries_len = self.entries.len() * 6;
        let complex_len: usize = self
            .entries
            .iter()
            .map(|e| match &e.value {
                FoptValue::Complex(bytes) => bytes.len(),
                FoptValue::Simple(_) => 0,
            })
            .sum();
        (entries_len + complex_len) as u32
    }

    /// Serialise the full FOPT atom (header + entries + complex
    /// trailer) into `out`.
    ///
    /// The `rec_instance` field of the header is set to the entry
    /// count, per MS-ODRAW §2.3.1; the `rec_ver` is 3 (atom).
    pub fn write_to(&self, out: &mut Vec<u8>) {
        let header = OfficeArtRecordHeader::atom(
            3,
            self.entries.len() as u16,
            rec_type::FOPT,
            self.payload_len(),
        );
        header.write_to(out);

        // Entry array.
        for entry in &self.entries {
            let mut opid = entry.id & 0x3FFF;
            if entry.is_blip_id {
                opid |= 0x4000;
            }
            let op = match &entry.value {
                FoptValue::Simple(v) => *v,
                FoptValue::Complex(bytes) => {
                    opid |= 0x8000;
                    bytes.len() as u32
                }
            };
            out.extend_from_slice(&opid.to_le_bytes());
            out.extend_from_slice(&op.to_le_bytes());
        }

        // Complex-data trailer.
        for entry in &self.entries {
            if let FoptValue::Complex(bytes) = &entry.value {
                out.extend_from_slice(bytes);
            }
        }
    }

    /// Parse a full FOPT atom (header included) from `bytes`.
    ///
    /// Returns `(table, bytes_consumed)`. The caller is expected to
    /// have positioned `bytes` at the start of the header.
    pub fn read_from(bytes: &[u8]) -> XlsResult<(Self, usize)> {
        let header = OfficeArtRecordHeader::read_from(bytes)?;
        if header.rec_type != rec_type::FOPT {
            return Err(XlsError::InvalidFormat(format!(
                "expected FOPT (0x{:04X}), found 0x{:04X}",
                rec_type::FOPT,
                header.rec_type
            )));
        }
        let entry_count = header.rec_instance as usize;
        let payload_start = HEADER_LEN;
        let entry_bytes_len = entry_count * 6;
        if bytes.len() < payload_start + entry_bytes_len {
            return Err(XlsError::InvalidFormat(format!(
                "FOPT truncated: need {} entry bytes, have {}",
                entry_bytes_len,
                bytes.len() - payload_start
            )));
        }

        // First pass: parse entry headers, remember complex-data lengths.
        let mut entries = Vec::with_capacity(entry_count);
        let mut complex_lens = Vec::with_capacity(entry_count);
        for i in 0..entry_count {
            let off = payload_start + i * 6;
            let opid = u16::from_le_bytes([bytes[off], bytes[off + 1]]);
            let op = u32::from_le_bytes([
                bytes[off + 2],
                bytes[off + 3],
                bytes[off + 4],
                bytes[off + 5],
            ]);
            let id = opid & 0x3FFF;
            let is_blip_id = (opid & 0x4000) != 0;
            let is_complex = (opid & 0x8000) != 0;
            if is_complex {
                complex_lens.push((entries.len(), op as usize));
                entries.push(FoptEntry {
                    id,
                    is_blip_id,
                    value: FoptValue::Complex(Vec::new()),
                });
            } else {
                entries.push(FoptEntry {
                    id,
                    is_blip_id,
                    value: FoptValue::Simple(op),
                });
            }
        }

        // Second pass: pull complex trailers in entry order.
        let mut cursor = payload_start + entry_bytes_len;
        let payload_end = payload_start + header.rec_len as usize;
        for (entry_idx, len) in complex_lens {
            if cursor + len > payload_end || cursor + len > bytes.len() {
                return Err(XlsError::InvalidFormat(format!(
                    "FOPT complex trailer for entry {entry_idx} overruns payload"
                )));
            }
            if let FoptValue::Complex(buf) = &mut entries[entry_idx].value {
                buf.extend_from_slice(&bytes[cursor..cursor + len]);
            }
            cursor += len;
        }

        Ok((Self { entries }, payload_start + header.rec_len as usize))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_dgg_container_header_round_trips() {
        // OfficeArtDggContainer with zero-length payload and zero instance.
        // Expected on-disk bytes: 0F 00 00 F0 00 00 00 00.
        let h = OfficeArtRecordHeader::container(rec_type::DGG_CONTAINER, 0, 0);
        let mut out = Vec::new();
        h.write_to(&mut out);
        assert_eq!(out, [0x0F, 0x00, 0x00, 0xF0, 0x00, 0x00, 0x00, 0x00]);

        let parsed = OfficeArtRecordHeader::read_from(&out).expect("header parses");
        assert_eq!(parsed, h);
        assert!(parsed.is_container());
    }

    #[test]
    fn fsp_atom_with_textbox_instance_encodes_in_high_nibbles() {
        // OfficeArtFSP recorded for a textbox shape: rec_ver=0x2,
        // rec_instance=msosptTextBox (0x00CA), rec_type=FSP (0xF00A),
        // rec_len=8 (FSP body is always 8 bytes).
        let h = OfficeArtRecordHeader::atom(0x2, shape_type::TEXT_BOX, rec_type::FSP, 8);
        let mut out = Vec::new();
        h.write_to(&mut out);
        // Byte 0: rec_ver=0x2 in low nibble, rec_instance low 4 bits
        // (0xCA & 0x0F = 0xA) in high nibble => 0xA2.
        // Byte 1: rec_instance bits 11..4 = 0x0C.
        assert_eq!(out, [0xA2, 0x0C, 0x0A, 0xF0, 0x08, 0x00, 0x00, 0x00]);

        let parsed = OfficeArtRecordHeader::read_from(&out).expect("FSP header parses");
        assert_eq!(parsed.rec_ver, 0x2);
        assert_eq!(parsed.rec_instance, shape_type::TEXT_BOX);
        assert_eq!(parsed.rec_type, rec_type::FSP);
        assert_eq!(parsed.rec_len, 8);
        assert!(!parsed.is_container());
    }

    #[test]
    fn split_menu_colors_uses_instance_4() {
        // SplitMenuColors always has rec_instance=4 (four ARGB entries),
        // rec_ver=0, rec_len=16.
        let h = OfficeArtRecordHeader::atom(0, 4, rec_type::SPLIT_MENU_COLORS, 16);
        let mut out = Vec::new();
        h.write_to(&mut out);
        // Byte 0: rec_ver=0, rec_instance low 4 bits = 4 => 0x40.
        // Byte 1: rec_instance bits 11..4 = 0.
        // Bytes 2-3: 0xF11E LE => 1E F1.
        // Bytes 4-7: 16 LE.
        assert_eq!(out, [0x40, 0x00, 0x1E, 0xF1, 0x10, 0x00, 0x00, 0x00]);

        let parsed = OfficeArtRecordHeader::read_from(&out).expect("split-menu header parses");
        assert_eq!(parsed.rec_instance, 4);
        assert_eq!(parsed.rec_type, rec_type::SPLIT_MENU_COLORS);
    }

    #[test]
    fn read_from_rejects_short_buffer() {
        let result = OfficeArtRecordHeader::read_from(&[0u8; 7]);
        match result {
            Err(XlsError::InvalidFormat(msg)) => {
                assert!(msg.contains("8"), "msg should mention 8 bytes: {msg}");
            }
            other => panic!("expected InvalidFormat error, got {other:?}"),
        }
    }

    #[test]
    fn container_helper_sets_rec_ver_to_f() {
        let h = OfficeArtRecordHeader::container(rec_type::SP_CONTAINER, 0, 100);
        assert_eq!(h.rec_ver, 0xF);
        assert!(h.is_container());
    }

    #[test]
    fn atom_helper_clamps_overflow_fields() {
        // rec_ver values > 0xF and rec_instance > 0xFFF must be
        // truncated, not panic. Caller may pass legitimate u16
        // shape-type values up to 0xFFF freely; the upper bits would
        // collide with rec_type's wire bytes.
        let h = OfficeArtRecordHeader::atom(0xFF, 0xFFFF, rec_type::FSP, 0);
        assert_eq!(h.rec_ver, 0x0F);
        assert_eq!(h.rec_instance, 0x0FFF);
    }

    #[test]
    fn total_len_includes_header() {
        let h = OfficeArtRecordHeader::atom(0, 0, rec_type::FOPT, 42);
        assert_eq!(h.total_len(), 8 + 42);
    }

    #[test]
    fn fopt_round_trips_three_simple_entries() {
        // A representative comment-textbox-ish FOPT:
        //   fillColor    = 0xFFFFE1 (light yellow)
        //   lineColor    = 0x000000
        //   group flags  = 0x000A0000 (fHidden + fPrint defaults)
        let mut t = FoptTable::new();
        t.push(FoptEntry::simple(fopt_id::FILL_COLOR, 0x00FF_FFE1));
        t.push(FoptEntry::simple(fopt_id::LINE_COLOR, 0));
        t.push(FoptEntry::simple(fopt_id::GROUP_SHAPE_PROPS, 0x000A_0000));

        let mut out = Vec::new();
        t.write_to(&mut out);

        // Header: rec_ver=3, rec_instance=3 (entry count), rec_type=FOPT
        // (0xF00B), rec_len=18 (3 entries × 6 bytes).
        // Byte 0: rec_ver=3 in low nibble, instance low 4 bits = 3 => 0x33.
        // Byte 1: instance bits 11..4 = 0 => 0x00.
        // Bytes 2-3: 0xF00B LE => 0B F0.
        // Bytes 4-7: 18 LE => 12 00 00 00.
        assert_eq!(&out[..8], &[0x33, 0x00, 0x0B, 0xF0, 0x12, 0x00, 0x00, 0x00]);

        let (parsed, consumed) = FoptTable::read_from(&out).expect("FOPT parses");
        assert_eq!(consumed, out.len());
        assert_eq!(parsed.len(), 3);
        let entries: Vec<_> = parsed.entries().cloned().collect();
        assert_eq!(
            entries[0],
            FoptEntry::simple(fopt_id::FILL_COLOR, 0x00FF_FFE1)
        );
        assert_eq!(entries[1], FoptEntry::simple(fopt_id::LINE_COLOR, 0));
        assert_eq!(
            entries[2],
            FoptEntry::simple(fopt_id::GROUP_SHAPE_PROPS, 0x000A_0000)
        );
    }

    #[test]
    fn fopt_complex_entry_trails_after_simple_array() {
        // One simple entry followed by one complex entry: the simple
        // value lives in the entry's op field, the complex bytes live
        // in the trailer.
        let mut t = FoptTable::new();
        t.push(FoptEntry::simple(fopt_id::FILL_COLOR, 0xDEADBEEF));
        t.push(FoptEntry::complex(0x017F, vec![0xAA, 0xBB, 0xCC, 0xDD]));

        let mut out = Vec::new();
        t.write_to(&mut out);

        // 8 header + 12 entries + 4 complex = 24 bytes total.
        assert_eq!(out.len(), 24);
        // The header's rec_len covers entries+complex (not header itself).
        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_len, 16);
        assert_eq!(header.rec_instance, 2);
        // Second entry's opid must have fComplex set: 0x017F | 0x8000 = 0x817F.
        let second_opid = u16::from_le_bytes([out[8 + 6], out[8 + 6 + 1]]);
        assert_eq!(second_opid, 0x817F);
        // And its op field carries the complex-data length, not the value.
        let second_op = u32::from_le_bytes([out[16], out[17], out[18], out[19]]);
        assert_eq!(second_op, 4);
        // Trailer bytes match what we pushed.
        assert_eq!(&out[20..24], &[0xAA, 0xBB, 0xCC, 0xDD]);

        let (parsed, _) = FoptTable::read_from(&out).unwrap();
        let entries: Vec<_> = parsed.entries().cloned().collect();
        assert_eq!(entries[0].id, fopt_id::FILL_COLOR);
        assert_eq!(entries[0].value, FoptValue::Simple(0xDEADBEEF));
        assert_eq!(entries[1].id, 0x017F);
        assert_eq!(
            entries[1].value,
            FoptValue::Complex(vec![0xAA, 0xBB, 0xCC, 0xDD])
        );
    }

    #[test]
    fn fopt_blip_id_entry_sets_fbid_bit() {
        // Picture-blip-reference entry: id=0x0104 (pibName? actually pib
        // is 0x0104), with fBid=1 and value=blip_id 1.
        let t = {
            let mut t = FoptTable::new();
            t.push(FoptEntry::blip_id(0x0104, 1));
            t
        };
        let mut out = Vec::new();
        t.write_to(&mut out);

        // opid = 0x0104 | 0x4000 (fBid) = 0x4104.
        let opid = u16::from_le_bytes([out[8], out[9]]);
        assert_eq!(opid, 0x4104);

        let (parsed, _) = FoptTable::read_from(&out).unwrap();
        let e = parsed.entries().next().unwrap();
        assert!(e.is_blip_id);
        assert_eq!(e.value, FoptValue::Simple(1));
    }

    #[test]
    fn fopt_read_rejects_wrong_rec_type() {
        // A header that claims to be DGG_CONTAINER instead of FOPT.
        let bogus = OfficeArtRecordHeader::container(rec_type::DGG_CONTAINER, 0, 0);
        let mut out = Vec::new();
        bogus.write_to(&mut out);
        let err = FoptTable::read_from(&out).unwrap_err();
        match err {
            XlsError::InvalidFormat(m) => assert!(m.contains("FOPT"), "msg: {m}"),
            other => panic!("expected InvalidFormat, got {other:?}"),
        }
    }

    #[test]
    fn fopt_round_trips_zero_entries() {
        let t = FoptTable::new();
        let mut out = Vec::new();
        t.write_to(&mut out);
        // 8-byte header + 0 payload.
        assert_eq!(out.len(), 8);
        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_instance, 0);
        assert_eq!(header.rec_len, 0);
        let (parsed, consumed) = FoptTable::read_from(&out).unwrap();
        assert_eq!(consumed, 8);
        assert!(parsed.is_empty());
    }

    #[test]
    fn fsp_round_trips_textbox_shape() {
        let fsp = OfficeArtFsp {
            spid: 0x0400, // shape id 1024
            grf_persistence: fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT,
        };
        let mut out = Vec::new();
        fsp.write_to(shape_type::TEXT_BOX, &mut out);

        // Header: rec_ver=2, rec_instance=TEXT_BOX (0xCA), rec_type=FSP, rec_len=8.
        // Byte 0: 2 (low) | (0xCA & 0x0F) << 4 = 2 | 0xA0 = 0xA2.
        // Byte 1: 0xCA >> 4 = 0x0C.
        // Then 0x0A 0xF0 (FSP type LE), 0x08 0x00 0x00 0x00 (len=8).
        assert_eq!(&out[..8], &[0xA2, 0x0C, 0x0A, 0xF0, 0x08, 0x00, 0x00, 0x00]);
        // Body: spid LE + flags LE.
        assert_eq!(&out[8..12], &0x0400u32.to_le_bytes());
        assert_eq!(
            &out[12..16],
            &(fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT).to_le_bytes()
        );

        let (parsed, st, consumed) = OfficeArtFsp::read_from(&out).unwrap();
        assert_eq!(parsed, fsp);
        assert_eq!(st, shape_type::TEXT_BOX);
        assert_eq!(consumed, out.len());
    }

    #[test]
    fn fsp_read_rejects_wrong_body_length() {
        // Build a deliberately corrupt FSP with rec_len=4.
        let mut out = Vec::new();
        let bad_header = OfficeArtRecordHeader::atom(2, shape_type::TEXT_BOX, rec_type::FSP, 4);
        bad_header.write_to(&mut out);
        out.extend_from_slice(&[0, 0, 0, 0]);
        let err = OfficeArtFsp::read_from(&out).unwrap_err();
        match err {
            XlsError::InvalidFormat(m) => assert!(m.contains("8 bytes"), "msg: {m}"),
            other => panic!("expected InvalidFormat, got {other:?}"),
        }
    }

    #[test]
    fn fspgr_round_trips_zero_rect() {
        let fspgr = OfficeArtFspgr::default();
        let mut out = Vec::new();
        fspgr.write_to(&mut out);

        // Header: rec_ver=1, rec_instance=0, rec_type=FSPGR (0xF009), rec_len=16.
        assert_eq!(&out[..8], &[0x01, 0x00, 0x09, 0xF0, 0x10, 0x00, 0x00, 0x00]);
        // Body: four LE i32 zeros.
        assert_eq!(&out[8..24], &[0u8; 16]);

        let (parsed, consumed) = OfficeArtFspgr::read_from(&out).unwrap();
        assert_eq!(parsed, fspgr);
        assert_eq!(consumed, 24);
    }

    #[test]
    fn fspgr_preserves_negative_coordinates() {
        let fspgr = OfficeArtFspgr {
            x_left: -100,
            y_top: -200,
            x_right: 300,
            y_bottom: 400,
        };
        let mut out = Vec::new();
        fspgr.write_to(&mut out);
        let (parsed, _) = OfficeArtFspgr::read_from(&out).unwrap();
        assert_eq!(parsed, fspgr);
    }

    #[test]
    fn client_anchor_round_trips_comment_default() {
        // Comment at row=10, col=4 should produce an anchor starting
        // at (col=5, row=10) on the same row, extending two columns
        // wide and four rows tall — matching Excel's own default
        // placement.
        let anchor = OfficeArtClientAnchor::comment_default(10, 4);
        assert_eq!(anchor.flag, 3);
        assert_eq!(anchor.col_l, 5);
        assert_eq!(anchor.row_t, 10);
        assert_eq!(anchor.col_r, 7);
        assert_eq!(anchor.row_b, 14);
        assert_eq!(anchor.dx_l, 240);
        assert_eq!(anchor.dy_t, 26);
        assert_eq!(anchor.dx_r, 496);
        assert_eq!(anchor.dy_b, 13);

        let mut out = Vec::new();
        anchor.write_to(&mut out);

        // Header: rec_ver=0, rec_instance=0, rec_type=ClientAnchor (0xF010), rec_len=18.
        assert_eq!(&out[..8], &[0x00, 0x00, 0x10, 0xF0, 0x12, 0x00, 0x00, 0x00]);

        let (parsed, consumed) = OfficeArtClientAnchor::read_from(&out).unwrap();
        assert_eq!(parsed, anchor);
        assert_eq!(consumed, 26); // 8 header + 18 body
    }

    #[test]
    fn client_data_writes_empty_atom() {
        let mut out = Vec::new();
        write_client_data(&mut out);
        // Header: rec_ver=0, rec_instance=0, rec_type=CLIENT_DATA (0xF011), rec_len=0.
        assert_eq!(out, [0x00, 0x00, 0x11, 0xF0, 0x00, 0x00, 0x00, 0x00]);
    }

    #[test]
    fn client_textbox_writes_empty_atom() {
        let mut out = Vec::new();
        write_client_textbox(&mut out);
        // Header: rec_ver=0, rec_instance=0, rec_type=CLIENT_TEXTBOX (0xF00D), rec_len=0.
        assert_eq!(out, [0x00, 0x00, 0x0D, 0xF0, 0x00, 0x00, 0x00, 0x00]);
    }

    #[test]
    fn client_anchor_handles_row_overflow_gracefully() {
        // Anchoring near u16::MAX must not panic; min() clamps row_t
        // at u16::MAX and saturating_add caps col_l and row_b.
        let anchor = OfficeArtClientAnchor::comment_default(u16::MAX as u32, u16::MAX - 5);
        assert_eq!(anchor.row_t, u16::MAX);
        assert_eq!(anchor.col_l, u16::MAX - 4);
    }

    #[test]
    fn fdgg_round_trips_one_cluster_one_drawing() {
        // Single sheet, single comment: spid_max=1025 (1024 base + 1),
        // one cluster for drawing 1 with one shape used.
        let f = OfficeArtFdgg {
            spid_max: 1025,
            csp_saved: 2, // patriarch + comment
            cdg_saved: 1,
            clusters: vec![IdCluster {
                dgid: 1,
                cspid_cur: 2,
            }],
        };
        let mut out = Vec::new();
        f.write_to(&mut out);

        // Header: rec_ver=0, rec_instance=0, rec_type=FDGG (0xF006),
        // rec_len = 16 + 8 = 24.
        assert_eq!(&out[..8], &[0x00, 0x00, 0x06, 0xF0, 0x18, 0x00, 0x00, 0x00]);
        // Body: spid_max=1025, cidcl=2 (clusters.len()+1), csp_saved=2,
        // cdg_saved=1, then 8 bytes for the cluster.
        assert_eq!(&out[8..12], &1025u32.to_le_bytes());
        assert_eq!(&out[12..16], &2u32.to_le_bytes());
        assert_eq!(&out[16..20], &2u32.to_le_bytes());
        assert_eq!(&out[20..24], &1u32.to_le_bytes());
        assert_eq!(&out[24..28], &1u32.to_le_bytes()); // cluster dgid
        assert_eq!(&out[28..32], &2u32.to_le_bytes()); // cluster cspid_cur

        let (parsed, consumed) = OfficeArtFdgg::read_from(&out).unwrap();
        assert_eq!(parsed, f);
        assert_eq!(consumed, out.len());
    }

    #[test]
    fn fdgg_round_trips_zero_clusters() {
        let f = OfficeArtFdgg {
            spid_max: 0,
            csp_saved: 0,
            cdg_saved: 0,
            clusters: vec![],
        };
        let mut out = Vec::new();
        f.write_to(&mut out);
        // cidcl is `clusters.len() + 1` = 1, even for an empty cluster table.
        let cidcl = u32::from_le_bytes([out[12], out[13], out[14], out[15]]);
        assert_eq!(cidcl, 1);
        let (parsed, _) = OfficeArtFdgg::read_from(&out).unwrap();
        assert_eq!(parsed, f);
    }

    #[test]
    fn fdg_round_trips_with_dgid_in_instance() {
        let f = OfficeArtFdg {
            csp_saved: 2,
            spid_last: 1025,
        };
        let mut out = Vec::new();
        f.write_to(1, &mut out);
        // Header: rec_ver=0, rec_instance=1 (dgid), rec_type=FDG (0xF008),
        // rec_len=8.
        // Byte 0: rec_ver=0 in low nibble, instance low 4 bits = 1 => 0x10.
        // Byte 1: instance bits 11..4 = 0.
        assert_eq!(&out[..8], &[0x10, 0x00, 0x08, 0xF0, 0x08, 0x00, 0x00, 0x00]);
        // Body: csp_saved=2, spid_last=1025.
        assert_eq!(&out[8..12], &2u32.to_le_bytes());
        assert_eq!(&out[12..16], &1025u32.to_le_bytes());

        let (parsed, dgid, consumed) = OfficeArtFdg::read_from(&out).unwrap();
        assert_eq!(parsed, f);
        assert_eq!(dgid, 1);
        assert_eq!(consumed, 16);
    }

    #[test]
    fn split_menu_colors_writes_excel_default_palette() {
        let mut out = Vec::new();
        SplitMenuColors::EXCEL_DEFAULT.write_to(&mut out);
        // 8 header + 16 body = 24 bytes.
        assert_eq!(out.len(), 24);
        // Header bytes match the existing header round-trip test.
        assert_eq!(&out[..8], &[0x40, 0x00, 0x1E, 0xF1, 0x10, 0x00, 0x00, 0x00]);
        // Body: the four scheme-colour entries Excel itself writes
        // (verified via the COM bridge probe).
        assert_eq!(&out[8..12], &0x0800_000Du32.to_le_bytes());
        assert_eq!(&out[12..16], &0x0800_000Cu32.to_le_bytes());
        assert_eq!(&out[16..20], &0x0800_0017u32.to_le_bytes());
        assert_eq!(&out[20..24], &0x1000_00F7u32.to_le_bytes());
    }

    #[test]
    fn write_container_wraps_body_with_correct_header() {
        // Empty SP_CONTAINER: rec_ver=0xF, rec_instance=0, rec_len=0.
        let mut out = Vec::new();
        write_container(rec_type::SP_CONTAINER, 0, &[], &mut out);
        assert_eq!(out, [0x0F, 0x00, 0x04, 0xF0, 0x00, 0x00, 0x00, 0x00]);

        // Container with three body bytes: rec_len=3.
        let mut out2 = Vec::new();
        write_container(rec_type::DG_CONTAINER, 1, &[0xAA, 0xBB, 0xCC], &mut out2);
        // Byte 0: rec_ver=0xF in low nibble, instance low 4 bits = 1 => 0x1F.
        // Byte 1: instance bits 11..4 = 0.
        assert_eq!(
            &out2[..8],
            &[0x1F, 0x00, 0x02, 0xF0, 0x03, 0x00, 0x00, 0x00]
        );
        assert_eq!(&out2[8..], &[0xAA, 0xBB, 0xCC]);
    }

    #[test]
    fn comment_sp_container_contains_expected_children() {
        let mut props = FoptTable::new();
        props.push(FoptEntry::simple(fopt_id::FILL_COLOR, 0x00FF_FFE1));
        props.push(FoptEntry::simple(fopt_id::GROUP_SHAPE_PROPS, 0));

        let mut out = Vec::new();
        write_comment_sp_container(1024, 3, 2, &props, &mut out);

        // The first 8 bytes must be an SP_CONTAINER header.
        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_type, rec_type::SP_CONTAINER);
        assert!(header.is_container());

        // Walk the body and confirm we see FSP, FOPT, ClientAnchor,
        // ClientData, ClientTextbox in that order.
        let mut cursor = HEADER_LEN;
        let saw: Vec<u16> = std::iter::from_fn(|| {
            if cursor >= out.len() {
                return None;
            }
            let h = OfficeArtRecordHeader::read_from(&out[cursor..]).unwrap();
            cursor += HEADER_LEN + h.rec_len as usize;
            Some(h.rec_type)
        })
        .collect();
        assert_eq!(
            saw,
            vec![
                rec_type::FSP,
                rec_type::FOPT,
                rec_type::CLIENT_ANCHOR,
                rec_type::CLIENT_DATA,
                rec_type::CLIENT_TEXTBOX,
            ]
        );

        // FSP's spid must be 1024.
        let (fsp, st, _) = OfficeArtFsp::read_from(&out[HEADER_LEN..]).unwrap();
        assert_eq!(fsp.spid, 1024);
        assert_eq!(st, shape_type::TEXT_BOX);
        assert_eq!(
            fsp.grf_persistence,
            fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT
        );
    }

    #[test]
    fn comment_fopt_has_14_entries_matching_excel() {
        // Verified by driving Excel via the bridge: every freshly-
        // created comment gets a 14-entry FOPT in this exact order.
        let t = comment_fopt(0x12345678);
        let entries: Vec<_> = t.entries().cloned().collect();
        assert_eq!(entries.len(), 14);
        // Spot-check the well-known IDs and values.
        assert_eq!(entries[0].id, fopt_id::TEXT_ID);
        assert_eq!(entries[0].value, FoptValue::Simple(0x12345678));
        assert_eq!(entries[2].id, fopt_id::TEXT_BOOLEAN_PROPS);
        assert_eq!(entries[2].value, FoptValue::Simple(0x0008_0008));
        assert_eq!(entries[4].id, fopt_id::FILL_COLOR);
        assert_eq!(entries[4].value, FoptValue::Simple(0x0800_0050));
        assert_eq!(entries[8].id, fopt_id::LINE_COLOR);
        assert_eq!(entries[8].value, FoptValue::Simple(0x0800_0051));
        assert_eq!(entries[13].id, fopt_id::GROUP_SHAPE_PROPS);
        assert_eq!(entries[13].value, FoptValue::Simple(0x0002_0002));

        // Serialise and confirm rec_instance carries the entry count.
        let mut out = Vec::new();
        t.write_to(&mut out);
        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_instance, 14);
        assert_eq!(header.rec_len, 14 * 6);
    }

    #[test]
    fn dgg_default_fopt_has_3_entries_matching_excel() {
        let t = dgg_default_fopt();
        let entries: Vec<_> = t.entries().cloned().collect();
        assert_eq!(entries.len(), 3);
        assert_eq!(entries[0].id, fopt_id::TEXT_BOOLEAN_PROPS);
        assert_eq!(entries[0].value, FoptValue::Simple(0x0008_0008));
        assert_eq!(entries[1].id, fopt_id::FILL_COLOR);
        assert_eq!(entries[1].value, FoptValue::Simple(0x0800_0041));
        assert_eq!(entries[2].id, fopt_id::LINE_COLOR);
        assert_eq!(entries[2].value, FoptValue::Simple(0x0800_0040));
    }

    #[test]
    fn patriarch_sp_container_has_fspgr_then_fsp() {
        let mut out = Vec::new();
        write_patriarch_sp_container(1024, &mut out);

        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_type, rec_type::SP_CONTAINER);

        // First child should be FSPGR.
        let first = OfficeArtRecordHeader::read_from(&out[HEADER_LEN..]).unwrap();
        assert_eq!(first.rec_type, rec_type::FSPGR);

        // Second child should be FSP with the PATRIARCH + GROUP flags.
        let second_off = HEADER_LEN + HEADER_LEN + first.rec_len as usize;
        let (fsp, st, _) = OfficeArtFsp::read_from(&out[second_off..]).unwrap();
        assert_eq!(fsp.spid, 1024);
        assert_eq!(st, shape_type::NOT_PRIMITIVE);
        assert_eq!(fsp.grf_persistence, fsp_flags::GROUP | fsp_flags::PATRIARCH);
    }

    #[test]
    fn blip_png_round_trips_with_md5_uid() {
        let png_bytes = vec![0xDE, 0xAD, 0xBE, 0xEF, 0x12, 0x34];
        let blip = OfficeArtBlip::png(png_bytes.clone());

        // UID must be deterministic across runs over the same input.
        assert_eq!(blip.rgb_uid, rgb_uid_for(&png_bytes));
        assert_ne!(
            blip.rgb_uid, [0u8; 16],
            "MD5 of non-empty bytes is non-zero"
        );

        let mut out = Vec::new();
        blip.write_to(&mut out);

        // Header: rec_ver=0, rec_instance=0x6E0, rec_type=0xF01E,
        // rec_len = 16 UID + 1 tag + N bytes.
        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_ver, 0);
        assert_eq!(header.rec_instance, blip_instance::PNG);
        assert_eq!(header.rec_type, rec_type::BLIP_PNG);
        assert_eq!(header.rec_len, (16 + 1 + png_bytes.len()) as u32);

        // UID + tag + data layout.
        assert_eq!(&out[HEADER_LEN..HEADER_LEN + 16], &blip.rgb_uid);
        assert_eq!(out[HEADER_LEN + 16], 0xFF);
        assert_eq!(&out[HEADER_LEN + 17..], &png_bytes[..]);
    }

    #[test]
    fn fbse_total_size_matches_serialised_bytes() {
        let blip = OfficeArtBlip::png(vec![0xAA; 50]);
        let blip_total = blip.total_size();
        let fbse = OfficeArtFbse::new(blip);

        let mut out = Vec::new();
        fbse.write_to(&mut out);

        // FBSE total = 8 (header) + 36 (fixed body) + embedded blip total.
        assert_eq!(out.len() as u32, 8 + 36 + blip_total);
        assert_eq!(fbse.total_size(), out.len() as u32);

        // First 8 bytes are the FBSE record header.
        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_ver, 2);
        assert_eq!(header.rec_instance, blip_type::PNG as u16);
        assert_eq!(header.rec_type, rec_type::FBSE);
        assert_eq!(header.rec_len, 36 + blip_total);

        // FBSE body byte 0 = btWin32 = PNG code (0x06).
        assert_eq!(out[HEADER_LEN], blip_type::PNG);
        // FBSE body byte 1 = btMacOS = PNG code.
        assert_eq!(out[HEADER_LEN + 1], blip_type::PNG);
        // Tag bytes at body offset 18-19 = 0xFF 0x00.
        assert_eq!(&out[HEADER_LEN + 18..HEADER_LEN + 20], &[0xFF, 0x00]);
        // size field at body offset 20-23 should equal blip_total.
        let size = u32::from_le_bytes([
            out[HEADER_LEN + 20],
            out[HEADER_LEN + 21],
            out[HEADER_LEN + 22],
            out[HEADER_LEN + 23],
        ]);
        assert_eq!(size, blip_total);
        // cRef = 1.
        let cref = u32::from_le_bytes([
            out[HEADER_LEN + 24],
            out[HEADER_LEN + 25],
            out[HEADER_LEN + 26],
            out[HEADER_LEN + 27],
        ]);
        assert_eq!(cref, 1);
    }

    #[test]
    fn bstore_container_wraps_one_fbse() {
        let blip = OfficeArtBlip::png(vec![0x01, 0x02, 0x03]);
        let fbse = OfficeArtFbse::new(blip);
        let fbse_total = fbse.total_size();

        let mut out = Vec::new();
        write_bstore_container(&[fbse], &mut out);

        // BSTORE_CONTAINER header: rec_ver=0xF, rec_instance=entry count = 1.
        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_type, rec_type::BSTORE_CONTAINER);
        assert!(header.is_container());
        assert_eq!(header.rec_instance, 1);
        assert_eq!(header.rec_len, fbse_total);

        // First inner record is an FBSE.
        let inner = OfficeArtRecordHeader::read_from(&out[HEADER_LEN..]).unwrap();
        assert_eq!(inner.rec_type, rec_type::FBSE);
    }

    #[test]
    fn picture_fopt_has_8_entries_and_wzname_complex_trailer() {
        let t = picture_fopt(1, "Picture 1");
        let entries: Vec<_> = t.entries().cloned().collect();
        assert_eq!(entries.len(), 8);

        // Entry 2 is the blip-id reference: opid=0x0104 with fBid bit set.
        assert_eq!(entries[1].id, 0x0104);
        assert!(entries[1].is_blip_id, "pib (0x0104) must set fBid");
        assert_eq!(entries[1].value, FoptValue::Simple(1));

        // Entry 7 is the wzName complex property containing "Picture 1\0"
        // in UTF-16LE, with both fBid and fComplex bits set (matching
        // Excel's emit pattern).
        assert_eq!(entries[6].id, 0x0380);
        assert!(entries[6].is_blip_id, "Excel sets fBid on wzName");
        match &entries[6].value {
            FoptValue::Complex(bytes) => {
                // 9 chars × 2 + 2 null bytes = 20 bytes.
                assert_eq!(bytes.len(), 20);
                // First two bytes = 'P' = 0x50 0x00 in UTF-16LE.
                assert_eq!(&bytes[..2], &[0x50, 0x00]);
            }
            FoptValue::Simple(_) => panic!("wzName must be a complex entry"),
        }
    }

    #[test]
    fn picture_sp_container_has_fsp_fopt_anchor_clientdata_no_textbox() {
        let mut out = Vec::new();
        let anchor = OfficeArtClientAnchor {
            flag: 2,
            col_l: 2,
            dx_l: 80,
            row_t: 6,
            dy_t: 166,
            col_r: 3,
            dx_r: 128,
            row_b: 10,
            dy_b: 0,
        };
        write_picture_sp_container(1025, 1, "Picture 1", anchor, &mut out);

        // Top-level: SP_CONTAINER.
        let header = OfficeArtRecordHeader::read_from(&out).unwrap();
        assert_eq!(header.rec_type, rec_type::SP_CONTAINER);
        assert!(header.is_container());

        // Walk children, confirm order: FSP, FOPT, ClientAnchor,
        // ClientData. No ClientTextbox (pictures have no text).
        let mut cursor = HEADER_LEN;
        let mut saw = Vec::new();
        while cursor < out.len() {
            let h = OfficeArtRecordHeader::read_from(&out[cursor..]).unwrap();
            saw.push(h.rec_type);
            cursor += HEADER_LEN + h.rec_len as usize;
        }
        assert_eq!(
            saw,
            vec![
                rec_type::FSP,
                rec_type::FOPT,
                rec_type::CLIENT_ANCHOR,
                rec_type::CLIENT_DATA,
            ]
        );

        // FSP must carry the picture-frame shape type and the supplied spid.
        let (fsp, st, _) = OfficeArtFsp::read_from(&out[HEADER_LEN..]).unwrap();
        assert_eq!(fsp.spid, 1025);
        assert_eq!(st, shape_type::PICTURE_FRAME);
        assert_eq!(
            fsp.grf_persistence,
            fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT
        );
    }
}
