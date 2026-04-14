// Wire format: b"DSBDL001" magic (8B) | entry_count (u32le)
//   per entry: path_len (u32le) | path (UTF-8) | data_len (u32le) | data

const MAGIC: &[u8; 8] = b"DSBDL001";

pub(crate) struct DrawingBundle {
    pub entries: Vec<(String, Vec<u8>)>,
}

impl DrawingBundle {
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    pub fn push(&mut self, path: String, data: Vec<u8>) {
        self.entries.push((path, data));
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn encode(&self) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(MAGIC);
        out.extend_from_slice(&(self.entries.len() as u32).to_le_bytes());
        for (path, data) in &self.entries {
            let path_bytes = path.as_bytes();
            out.extend_from_slice(&(path_bytes.len() as u32).to_le_bytes());
            out.extend_from_slice(path_bytes);
            out.extend_from_slice(&(data.len() as u32).to_le_bytes());
            out.extend_from_slice(data);
        }
        out
    }

    pub fn decode(bytes: &[u8]) -> Option<Self> {
        if bytes.len() < 12 || &bytes[..8] != MAGIC {
            return None;
        }
        let count = u32::from_le_bytes(bytes[8..12].try_into().ok()?) as usize;
        let mut pos = 12;
        let mut entries = Vec::with_capacity(count);
        for _ in 0..count {
            if pos + 4 > bytes.len() {
                return None;
            }
            let path_len = u32::from_le_bytes(bytes[pos..pos + 4].try_into().ok()?) as usize;
            pos += 4;
            if pos + path_len > bytes.len() {
                return None;
            }
            let path = std::str::from_utf8(&bytes[pos..pos + path_len])
                .ok()?
                .to_string();
            pos += path_len;
            if pos + 4 > bytes.len() {
                return None;
            }
            let data_len = u32::from_le_bytes(bytes[pos..pos + 4].try_into().ok()?) as usize;
            pos += 4;
            if pos + data_len > bytes.len() {
                return None;
            }
            let data = bytes[pos..pos + data_len].to_vec();
            pos += data_len;
            entries.push((path, data));
        }
        Some(Self { entries })
    }
}

pub(crate) fn is_drawing_bundle(bytes: &[u8]) -> bool {
    bytes.len() >= 8 && &bytes[..8] == MAGIC
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn roundtrip_bundle() {
        let mut b = DrawingBundle::new();
        b.push("xl/drawings/drawing1.xml".into(), b"<xdr:wsDr/>".to_vec());
        b.push("xl/charts/chart1.xml".into(), b"<c:chartSpace/>".to_vec());

        let encoded = b.encode();
        assert!(is_drawing_bundle(&encoded));

        let decoded = DrawingBundle::decode(&encoded).unwrap();
        assert_eq!(decoded.entries.len(), 2);
        assert_eq!(decoded.entries[0].0, "xl/drawings/drawing1.xml");
        assert_eq!(decoded.entries[0].1, b"<xdr:wsDr/>");
        assert_eq!(decoded.entries[1].0, "xl/charts/chart1.xml");
        assert_eq!(decoded.entries[1].1, b"<c:chartSpace/>");
    }

    #[test]
    fn empty_bundle() {
        let b = DrawingBundle::new();
        assert!(b.is_empty());
        let encoded = b.encode();
        let decoded = DrawingBundle::decode(&encoded).unwrap();
        assert!(decoded.is_empty());
    }

    #[test]
    fn not_a_bundle() {
        assert!(!is_drawing_bundle(b"<xdr:wsDr/>"));
        assert!(!is_drawing_bundle(b"short"));
        assert!(DrawingBundle::decode(b"notabundle").is_none());
    }
}
