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
    /// Build a "comment-default" anchor that places the textbox
    /// starting just past the right edge of `(row, col)` and spanning
    /// roughly two cells wide × four rows tall — the layout Excel
    /// itself uses for newly-created comments. The exact pixel
    /// offsets match the BIFF8 defaults observed in Excel output.
    pub fn comment_default(row: u32, col: u16) -> Self {
        // Excel's defaults: start one column right and one row down,
        // extend two columns and four rows further.
        let row_t = row.saturating_add(1).min(u16::MAX as u32) as u16;
        let col_l = col.saturating_add(1);
        Self {
            flag: 0,
            col_l,
            dx_l: 15 * 64, // ~15/64ths into the cell
            row_t,
            dy_t: 0,
            col_r: col_l.saturating_add(2),
            dx_r: 15 * 64,
            row_b: row_t.saturating_add(4),
            dy_b: 0,
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
        // at (col=5, row=11) and extending two columns + four rows.
        let anchor = OfficeArtClientAnchor::comment_default(10, 4);
        assert_eq!(anchor.col_l, 5);
        assert_eq!(anchor.row_t, 11);
        assert_eq!(anchor.col_r, 7);
        assert_eq!(anchor.row_b, 15);

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
        // Anchoring near u16::MAX must not panic. saturating_add caps
        // row_t at u16::MAX; row_b clamps relative to that.
        let anchor = OfficeArtClientAnchor::comment_default(u16::MAX as u32, u16::MAX - 5);
        assert_eq!(anchor.row_t, u16::MAX);
        assert_eq!(anchor.col_l, u16::MAX - 4);
    }
}
