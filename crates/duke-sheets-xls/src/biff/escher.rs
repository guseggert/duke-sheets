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
}
