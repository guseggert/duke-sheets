use std::fmt;
use std::io::{self, Read, Seek};

const HEADER_LEN: usize = 512;
const DIR_ENTRY_LEN: usize = 128;

const CFB_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

const MAX_REGULAR_SECTOR: u32 = 0xFFFFFFF9;

const FREESECT: u32 = 0xFFFFFFFF;
const ENDOFCHAIN: u32 = 0xFFFFFFFE;
const NOSTREAM: u32 = 0xFFFFFFFF;

#[derive(Debug)]
pub enum CfbError {
    Io(io::Error),
    InvalidFormat(String),
}

impl fmt::Display for CfbError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CfbError::Io(err) => write!(f, "{err}"),
            CfbError::InvalidFormat(msg) => write!(f, "invalid CFB format: {msg}"),
        }
    }
}

impl From<io::Error> for CfbError {
    fn from(value: io::Error) -> Self {
        CfbError::Io(value)
    }
}

impl From<CfbError> for io::Error {
    fn from(value: CfbError) -> Self {
        match value {
            CfbError::Io(err) => err,
            CfbError::InvalidFormat(msg) => io::Error::new(io::ErrorKind::InvalidData, msg),
        }
    }
}

#[derive(Clone, Debug)]
pub struct DirectoryEntry {
    name: String,
    object_type: u8,
    // Color is parsed for RB-invariant assertions in tests but not
    // consulted by production read paths (we walk the tree as a plain
    // BST). Suppress dead-code warning in non-test builds.
    #[cfg_attr(not(test), allow(dead_code))]
    color: u8,
    left_sibling: u32,
    right_sibling: u32,
    child: u32,
    start_sector: u32,
    stream_size: u64,
}

impl DirectoryEntry {
    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn object_type(&self) -> u8 {
        self.object_type
    }

    pub fn stream_size(&self) -> u64 {
        self.stream_size
    }
}

pub struct CompoundFile {
    file_data: Vec<u8>,
    sector_size: usize,
    mini_sector_size: usize,
    mini_stream_cutoff: usize,
    fat: Vec<u32>,
    mini_fat: Vec<u32>,
    directory: Vec<DirectoryEntry>,
    mini_stream_data: Vec<u8>,
}

impl CompoundFile {
    pub fn open<R: Read + Seek>(mut reader: R) -> Result<Self, CfbError> {
        let mut file_data = Vec::new();
        reader.read_to_end(&mut file_data)?;

        if file_data.len() < HEADER_LEN {
            return Err(CfbError::InvalidFormat(
                "file too small for CFB header".into(),
            ));
        }

        if file_data[0..8] != CFB_MAGIC {
            return Err(CfbError::InvalidFormat(
                "missing CFB magic signature".into(),
            ));
        }

        let sector_shift = read_u16(&file_data, 30)?;
        let mini_sector_shift = read_u16(&file_data, 32)?;
        let sector_size = checked_pow2(sector_shift)
            .ok_or_else(|| CfbError::InvalidFormat("invalid sector size shift".into()))?;
        let mini_sector_size = checked_pow2(mini_sector_shift)
            .ok_or_else(|| CfbError::InvalidFormat("invalid mini sector size shift".into()))?;

        if sector_size != 512 && sector_size != 4096 {
            return Err(CfbError::InvalidFormat(format!(
                "unsupported sector size: {sector_size}"
            )));
        }

        if mini_sector_size != 64 {
            return Err(CfbError::InvalidFormat(format!(
                "unsupported mini sector size: {mini_sector_size}"
            )));
        }

        // Some third-party XLS writers leave the final sector truncated on disk
        // rather than zero-padding it to the sector boundary the spec requires.
        // Pad in memory so `sector_offset` sees a complete file.
        pad_to_sector_boundary(&mut file_data, sector_size);

        let total_fat_sectors = read_u32(&file_data, 44)? as usize;

        // Sanity check: header sector counts can't exceed the number of
        // regular sectors the file can physically contain. Sector id 0
        // starts immediately after the header, so subtract the header slot.
        let max_regular_sectors = regular_sector_count(file_data.len(), sector_size);
        if total_fat_sectors > max_regular_sectors {
            return Err(CfbError::InvalidFormat(format!(
                "total FAT sectors ({total_fat_sectors}) exceeds file capacity ({max_regular_sectors} sectors)"
            )));
        }
        let first_directory_sector = read_u32(&file_data, 48)?;
        let mini_stream_cutoff = read_u32(&file_data, 56)? as usize;
        let first_mini_fat_sector = read_u32(&file_data, 60)?;
        let total_mini_fat_sectors = read_u32(&file_data, 64)? as usize;
        let first_difat_sector = read_u32(&file_data, 68)?;
        let total_difat_sectors = read_u32(&file_data, 72)? as usize;

        if total_difat_sectors > max_regular_sectors {
            return Err(CfbError::InvalidFormat(format!(
                "total DIFAT sectors ({total_difat_sectors}) exceeds file capacity ({max_regular_sectors} sectors)"
            )));
        }
        if total_mini_fat_sectors > max_regular_sectors {
            return Err(CfbError::InvalidFormat(format!(
                "total mini FAT sectors ({total_mini_fat_sectors}) exceeds file capacity ({max_regular_sectors} sectors)"
            )));
        }

        let difat = read_difat(
            &file_data,
            sector_size,
            total_fat_sectors,
            first_difat_sector,
            total_difat_sectors,
        )?;

        let fat = read_fat(&file_data, sector_size, &difat, total_fat_sectors)?;

        let directory_stream =
            read_regular_chain(&file_data, sector_size, &fat, first_directory_sector, None)?;
        let directory = parse_directory(&directory_stream, sector_size)?;
        if directory.is_empty() {
            return Err(CfbError::InvalidFormat("empty directory stream".into()));
        }
        if directory[0].object_type != 5 {
            return Err(CfbError::InvalidFormat(
                "directory entry 0 is not Root Entry".into(),
            ));
        }

        let mini_fat = if total_mini_fat_sectors == 0 {
            Vec::new()
        } else {
            read_mini_fat(
                &file_data,
                sector_size,
                &fat,
                first_mini_fat_sector,
                total_mini_fat_sectors,
            )?
        };

        let root = &directory[0];
        let root_size = usize::try_from(root.stream_size).map_err(|_| {
            CfbError::InvalidFormat("root entry stream size does not fit in memory".into())
        })?;
        let mut mini_stream_data = if root_size == 0 || root.start_sector == ENDOFCHAIN {
            Vec::new()
        } else {
            read_regular_chain(
                &file_data,
                sector_size,
                &fat,
                root.start_sector,
                Some(root_size),
            )?
        };
        // Mirror the file-level padding: the mini stream is a byte-for-byte
        // concatenation of mini sectors, and some producers leave the root
        // stream size (root_size) shy of the actual mini-sector boundary.
        // Padding with zeros lets the mini FAT chain's final sector fall
        // within the buffer. Any real data in that tail would be zeros anyway
        // per spec.
        pad_to_sector_boundary(&mut mini_stream_data, mini_sector_size);

        Ok(Self {
            file_data,
            sector_size,
            mini_sector_size,
            mini_stream_cutoff,
            fat,
            mini_fat,
            directory,
            mini_stream_data,
        })
    }

    pub fn exists(&self, path: &str) -> bool {
        self.find_entry(path)
            .is_some_and(|entry| entry.object_type == 2)
    }

    /// Iterate over all directory entries (in storage order). Useful
    /// for tooling that needs to inspect the CFB structure beyond the
    /// `exists`/`read_stream` API.
    pub fn directory_entries(&self) -> impl Iterator<Item = &DirectoryEntry> {
        self.directory.iter()
    }

    pub fn read_stream(&self, path: &str) -> Result<Vec<u8>, CfbError> {
        let entry = self
            .find_entry(path)
            .ok_or_else(|| CfbError::InvalidFormat(format!("stream not found: {path}")))?;

        if entry.object_type != 2 {
            return Err(CfbError::InvalidFormat(format!(
                "path is not a stream: {path}"
            )));
        }

        let stream_size = usize::try_from(entry.stream_size)
            .map_err(|_| CfbError::InvalidFormat("stream too large to fit in memory".into()))?;

        if stream_size == 0 {
            return Ok(Vec::new());
        }

        if entry.start_sector == ENDOFCHAIN {
            return Err(CfbError::InvalidFormat(format!(
                "stream has ENDOFCHAIN start sector: {path}"
            )));
        }

        if stream_size < self.mini_stream_cutoff {
            self.read_mini_stream(entry.start_sector, stream_size)
        } else {
            read_regular_chain(
                &self.file_data,
                self.sector_size,
                &self.fat,
                entry.start_sector,
                Some(stream_size),
            )
        }
    }

    fn read_mini_stream(&self, start_mini_sector: u32, size: usize) -> Result<Vec<u8>, CfbError> {
        let mut out = Vec::with_capacity(size);
        let mut current = start_mini_sector;

        loop {
            if out.len() >= size {
                break;
            }
            if current == ENDOFCHAIN {
                break;
            }
            if current > MAX_REGULAR_SECTOR {
                return Err(CfbError::InvalidFormat(format!(
                    "invalid mini sector id: {current:#010X}"
                )));
            }

            let index = current as usize;
            if index >= self.mini_fat.len() {
                return Err(CfbError::InvalidFormat(
                    "mini FAT index out of range while reading stream".into(),
                ));
            }

            let offset = index
                .checked_mul(self.mini_sector_size)
                .ok_or_else(|| CfbError::InvalidFormat("mini stream offset overflow".into()))?;
            let end = offset
                .checked_add(self.mini_sector_size)
                .ok_or_else(|| CfbError::InvalidFormat("mini stream end overflow".into()))?;
            if end > self.mini_stream_data.len() {
                return Err(CfbError::InvalidFormat(
                    "mini stream sector points outside root mini stream".into(),
                ));
            }

            let remaining = size - out.len();
            let take = remaining.min(self.mini_sector_size);
            out.extend_from_slice(&self.mini_stream_data[offset..offset + take]);

            current = self.mini_fat[index];
        }

        if out.len() < size {
            return Err(CfbError::InvalidFormat(
                "mini stream ended before requested size".into(),
            ));
        }

        Ok(out)
    }

    fn find_entry(&self, path: &str) -> Option<&DirectoryEntry> {
        let mut segments = path.split('/').filter(|segment| !segment.is_empty());
        let first = segments.next()?;

        let mut current_id = self.find_child_by_name(0, first)?;

        for segment in segments {
            let current = self.directory.get(current_id)?;
            if current.object_type != 1 && current.object_type != 5 {
                return None;
            }
            current_id = self.find_child_by_name(current_id, segment)?;
        }

        self.directory.get(current_id)
    }

    fn find_child_by_name(&self, parent_id: usize, name: &str) -> Option<usize> {
        let parent = self.directory.get(parent_id)?;
        let root = parent.child;
        if root == NOSTREAM {
            return None;
        }
        if root > MAX_REGULAR_SECTOR {
            return None;
        }

        let mut stack = vec![root as usize];
        let mut visited = vec![false; self.directory.len()];

        while let Some(entry_id) = stack.pop() {
            if entry_id >= self.directory.len() || visited[entry_id] {
                continue;
            }
            visited[entry_id] = true;

            let entry = &self.directory[entry_id];
            // CFB stream names are technically case-sensitive in storage, but
            // in practice Windows and Office apps treat them case-insensitively
            // because Windows filesystem semantics do. Some tools write the
            // workbook stream as `WORKBOOK` (all caps) rather than Excel's
            // canonical `Workbook`; matching case-insensitively lets us read
            // those files.
            if entry.name.eq_ignore_ascii_case(name) {
                return Some(entry_id);
            }

            if entry.left_sibling <= MAX_REGULAR_SECTOR {
                stack.push(entry.left_sibling as usize);
            }
            if entry.right_sibling <= MAX_REGULAR_SECTOR {
                stack.push(entry.right_sibling as usize);
            }
        }

        None
    }
}

fn checked_pow2(shift: u16) -> Option<usize> {
    if shift >= usize::BITS as u16 {
        return None;
    }
    Some(1usize << shift)
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16, CfbError> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| CfbError::InvalidFormat("u16 offset overflow".into()))?;
    if end > data.len() {
        return Err(CfbError::InvalidFormat(
            "unexpected EOF while reading u16".into(),
        ));
    }
    Ok(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, CfbError> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| CfbError::InvalidFormat("u32 offset overflow".into()))?;
    if end > data.len() {
        return Err(CfbError::InvalidFormat(
            "unexpected EOF while reading u32".into(),
        ));
    }
    Ok(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
}

fn read_u64(data: &[u8], offset: usize) -> Result<u64, CfbError> {
    let end = offset
        .checked_add(8)
        .ok_or_else(|| CfbError::InvalidFormat("u64 offset overflow".into()))?;
    if end > data.len() {
        return Err(CfbError::InvalidFormat(
            "unexpected EOF while reading u64".into(),
        ));
    }
    Ok(u64::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]))
}

fn sector_offset(
    file_data_len: usize,
    sector_size: usize,
    sector_id: u32,
) -> Result<usize, CfbError> {
    if sector_id > MAX_REGULAR_SECTOR {
        return Err(CfbError::InvalidFormat(format!(
            "invalid regular sector id: {sector_id:#010X}"
        )));
    }

    let sector_index = sector_id as usize;
    let bytes = sector_index
        .checked_add(1)
        .and_then(|v| v.checked_mul(sector_size))
        .ok_or_else(|| CfbError::InvalidFormat("sector offset overflow".into()))?;

    let end = bytes
        .checked_add(sector_size)
        .ok_or_else(|| CfbError::InvalidFormat("sector end overflow".into()))?;

    if end > file_data_len {
        return Err(CfbError::InvalidFormat(
            "sector points outside file bounds".into(),
        ));
    }

    Ok(bytes)
}

fn regular_sector_count(file_data_len: usize, sector_size: usize) -> usize {
    (file_data_len / sector_size).saturating_sub(1)
}

fn read_difat(
    file_data: &[u8],
    sector_size: usize,
    total_fat_sectors: usize,
    first_difat_sector: u32,
    total_difat_sectors: usize,
) -> Result<Vec<u32>, CfbError> {
    if total_fat_sectors == 0 {
        return Ok(Vec::new());
    }

    let mut difat = Vec::with_capacity(total_fat_sectors);

    for i in 0..109usize {
        let offset = 76 + i * 4;
        let sid = read_u32(file_data, offset)?;
        if sid != FREESECT {
            difat.push(sid);
            if difat.len() == total_fat_sectors {
                return Ok(difat);
            }
        }
    }

    let mut current = first_difat_sector;
    let entries_per_difat_sector = sector_size / 4;
    if entries_per_difat_sector < 2 {
        return Err(CfbError::InvalidFormat(
            "sector size too small for DIFAT sector".into(),
        ));
    }

    for _ in 0..total_difat_sectors {
        if current == ENDOFCHAIN {
            break;
        }

        let offset = sector_offset(file_data.len(), sector_size, current)?;
        let sector = &file_data[offset..offset + sector_size];

        for i in 0..(entries_per_difat_sector - 1) {
            let sid = read_u32(sector, i * 4)?;
            if sid != FREESECT {
                difat.push(sid);
                if difat.len() == total_fat_sectors {
                    return Ok(difat);
                }
            }
        }

        current = read_u32(sector, (entries_per_difat_sector - 1) * 4)?;
    }

    if difat.len() < total_fat_sectors {
        return Err(CfbError::InvalidFormat(format!(
            "insufficient DIFAT entries: expected at least {total_fat_sectors}, got {}",
            difat.len()
        )));
    }

    difat.truncate(total_fat_sectors);
    Ok(difat)
}

fn read_fat(
    file_data: &[u8],
    sector_size: usize,
    difat: &[u32],
    total_fat_sectors: usize,
) -> Result<Vec<u32>, CfbError> {
    let entries_per_sector = sector_size / 4;
    let mut fat = Vec::with_capacity(total_fat_sectors * entries_per_sector);

    for &sector_id in difat.iter().take(total_fat_sectors) {
        let offset = sector_offset(file_data.len(), sector_size, sector_id)?;
        let sector = &file_data[offset..offset + sector_size];
        for i in 0..entries_per_sector {
            fat.push(read_u32(sector, i * 4)?);
        }
    }

    Ok(fat)
}

fn read_regular_chain(
    file_data: &[u8],
    sector_size: usize,
    fat: &[u32],
    start_sector: u32,
    expected_size: Option<usize>,
) -> Result<Vec<u8>, CfbError> {
    if start_sector == ENDOFCHAIN {
        return Ok(Vec::new());
    }
    if start_sector > MAX_REGULAR_SECTOR {
        return Err(CfbError::InvalidFormat(format!(
            "invalid start sector id: {start_sector:#010X}"
        )));
    }

    let mut out = Vec::new();
    if let Some(expected_size) = expected_size {
        let max_chain_bytes = regular_sector_count(file_data.len(), sector_size)
            .checked_mul(sector_size)
            .ok_or_else(|| CfbError::InvalidFormat("maximum chain size overflow".into()))?;
        if expected_size > max_chain_bytes {
            return Err(CfbError::InvalidFormat(format!(
                "declared stream size ({expected_size}) exceeds file capacity ({max_chain_bytes} bytes)"
            )));
        }
        out.reserve(expected_size);
    }

    let mut current = start_sector;
    let mut visited = vec![false; fat.len()];

    loop {
        if current == ENDOFCHAIN {
            break;
        }
        if current > MAX_REGULAR_SECTOR {
            return Err(CfbError::InvalidFormat(format!(
                "invalid chain sector id: {current:#010X}"
            )));
        }

        let index = current as usize;
        if index >= fat.len() {
            return Err(CfbError::InvalidFormat(
                "FAT index out of range while reading chain".into(),
            ));
        }
        if visited[index] {
            return Err(CfbError::InvalidFormat("loop detected in FAT chain".into()));
        }
        visited[index] = true;

        let offset = sector_offset(file_data.len(), sector_size, current)?;
        out.extend_from_slice(&file_data[offset..offset + sector_size]);

        current = fat[index];

        if let Some(expected_size) = expected_size {
            if out.len() >= expected_size {
                break;
            }
        }
    }

    if let Some(expected_size) = expected_size {
        if out.len() < expected_size {
            return Err(CfbError::InvalidFormat(
                "stream chain shorter than declared stream size".into(),
            ));
        }
        out.truncate(expected_size);
    }

    Ok(out)
}

fn read_mini_fat(
    file_data: &[u8],
    sector_size: usize,
    fat: &[u32],
    first_mini_fat_sector: u32,
    total_mini_fat_sectors: usize,
) -> Result<Vec<u32>, CfbError> {
    let raw = read_regular_chain(
        file_data,
        sector_size,
        fat,
        first_mini_fat_sector,
        Some(
            total_mini_fat_sectors
                .checked_mul(sector_size)
                .ok_or_else(|| CfbError::InvalidFormat("mini FAT size overflow".into()))?,
        ),
    )?;

    let entries = raw.len() / 4;
    let mut mini_fat = Vec::with_capacity(entries);
    for i in 0..entries {
        mini_fat.push(read_u32(&raw, i * 4)?);
    }
    Ok(mini_fat)
}

fn parse_directory(data: &[u8], sector_size: usize) -> Result<Vec<DirectoryEntry>, CfbError> {
    if data.len() < DIR_ENTRY_LEN {
        return Err(CfbError::InvalidFormat("directory stream too short".into()));
    }

    // MS-CFB §2.6.1: On CFB v3 (512-byte sectors) the StreamSize high DWORD is reserved
    // and "should be ignored" by consumers. Some older implementations (notably Excel 97)
    // leave garbage in those bytes, which would otherwise be read as a multi-petabyte
    // stream size and abort the allocator.
    let stream_size_is_u32 = sector_size == 512;

    let mut entries = Vec::with_capacity(data.len() / DIR_ENTRY_LEN);
    let mut offset = 0usize;
    while offset + DIR_ENTRY_LEN <= data.len() {
        let chunk = &data[offset..offset + DIR_ENTRY_LEN];

        let name_len_bytes = read_u16(chunk, 64)? as usize;
        if name_len_bytes > 64 || name_len_bytes % 2 != 0 {
            return Err(CfbError::InvalidFormat(
                "invalid directory name length".into(),
            ));
        }

        let char_count = if name_len_bytes >= 2 {
            (name_len_bytes / 2).saturating_sub(1)
        } else {
            0
        };

        let mut utf16 = Vec::with_capacity(char_count);
        for i in 0..char_count {
            let lo = chunk[i * 2];
            let hi = chunk[i * 2 + 1];
            utf16.push(u16::from_le_bytes([lo, hi]));
        }
        let name = String::from_utf16_lossy(&utf16);

        let raw_stream_size = read_u64(chunk, 120)?;
        let stream_size = if stream_size_is_u32 {
            raw_stream_size & 0xFFFF_FFFF
        } else {
            raw_stream_size
        };

        entries.push(DirectoryEntry {
            name,
            object_type: chunk[66],
            color: chunk[67],
            left_sibling: read_u32(chunk, 68)?,
            right_sibling: read_u32(chunk, 72)?,
            child: read_u32(chunk, 76)?,
            start_sector: read_u32(chunk, 116)?,
            stream_size,
        });

        offset += DIR_ENTRY_LEN;
    }

    Ok(entries)
}

/// Pad a CFB file buffer to the next sector boundary with zeros.
///
/// MS-CFB specifies that a compound file's byte length be a multiple of its
/// sector size, with any unused tail bytes zeroed. Some real-world XLS writers
/// skip that padding, so the last sector is truncated on disk and our
/// strict `sector_offset` check rejects the file. Padding with zeros at load
/// time matches what the producer should have written and lets us read valid
/// streams whose on-disk tails happened to be all zeros anyway.
fn pad_to_sector_boundary(file_data: &mut Vec<u8>, sector_size: usize) {
    let rem = file_data.len() % sector_size;
    if rem != 0 {
        let pad = sector_size - rem;
        file_data.resize(file_data.len() + pad, 0);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pad_to_sector_boundary_adds_zeros_when_truncated() {
        let mut data = vec![0xAB; 1500]; // 1500 bytes, sector_size 512 -> pad to 1536
        pad_to_sector_boundary(&mut data, 512);
        assert_eq!(data.len(), 1536);
        assert_eq!(&data[..1500], &vec![0xAB; 1500][..]);
        assert_eq!(&data[1500..], &[0; 36]);
    }

    #[test]
    fn pad_to_sector_boundary_is_noop_when_aligned() {
        let mut data = vec![0xAB; 1024];
        pad_to_sector_boundary(&mut data, 512);
        assert_eq!(data.len(), 1024);
    }

    #[test]
    fn pad_to_sector_boundary_handles_4k_sectors() {
        let mut data = vec![0xAB; 5000];
        pad_to_sector_boundary(&mut data, 4096);
        assert_eq!(data.len(), 8192);
    }

    #[test]
    fn rejects_impossible_difat_sector_count_before_allocation() {
        let mut data = vec![0u8; 545];
        data[..8].copy_from_slice(&CFB_MAGIC);
        data[30..32].copy_from_slice(&9u16.to_le_bytes()); // 512-byte sectors
        data[32..34].copy_from_slice(&6u16.to_le_bytes()); // 64-byte mini sectors
        data[44..48].copy_from_slice(&0u32.to_le_bytes()); // no FAT sectors needed
        data[68..72].copy_from_slice(&0u32.to_le_bytes()); // first DIFAT sector
        data[72..76].copy_from_slice(&0x9B00_00D9u32.to_le_bytes());

        let err = match CompoundFile::open(std::io::Cursor::new(data)) {
            Ok(_) => panic!("malformed CFB unexpectedly opened"),
            Err(err) => err,
        };
        assert!(
            err.to_string().contains("total DIFAT sectors"),
            "unexpected error: {err}"
        );
    }

    fn make_dir_entry(name: &str, object_type: u8, stream_size_raw: u64) -> [u8; DIR_ENTRY_LEN] {
        let mut e = [0u8; DIR_ENTRY_LEN];
        let mut utf16: Vec<u16> = name.encode_utf16().collect();
        utf16.push(0);
        for (i, w) in utf16.iter().enumerate() {
            let [lo, hi] = w.to_le_bytes();
            e[i * 2] = lo;
            e[i * 2 + 1] = hi;
        }
        let name_len_bytes = (utf16.len() * 2) as u16;
        e[64..66].copy_from_slice(&name_len_bytes.to_le_bytes());
        e[66] = object_type;
        e[68..72].copy_from_slice(&NOSTREAM.to_le_bytes());
        e[72..76].copy_from_slice(&NOSTREAM.to_le_bytes());
        e[76..80].copy_from_slice(&NOSTREAM.to_le_bytes());
        e[116..120].copy_from_slice(&0u32.to_le_bytes());
        e[120..128].copy_from_slice(&stream_size_raw.to_le_bytes());
        e
    }

    /// MS-CFB §2.6.1: "For a version 3 compound file 512-byte sector size, the value of
    /// this field MUST be less than or equal to 0x80000000. (Note: Some older
    /// implementations may have set the high DWORD of this field to something other than
    /// zero, in which case it should be ignored.)"
    ///
    /// Old Excel 97 files in the wild have garbage in the high DWORD. If we read the raw
    /// u64 and feed it to Vec::with_capacity, the allocator aborts the process.
    #[test]
    fn parse_directory_masks_v3_stream_size_high_dword() {
        let raw = 0x7000_4000_0000_B74Du64; // high garbage, low = 46925
        let mut data = Vec::new();
        data.extend_from_slice(&make_dir_entry("Root Entry", 5, 0));
        data.extend_from_slice(&make_dir_entry("Workbook", 2, raw));

        let entries = parse_directory(&data, 512).expect("parse ok");
        assert_eq!(entries.len(), 2);
        assert_eq!(
            entries[1].stream_size, 46925,
            "v3 stream_size must be masked to low 32 bits (got {:#x})",
            entries[1].stream_size
        );
    }

    /// CFB v4 (4096-byte sector) uses the full u64 StreamSize field, so it must not be
    /// masked.
    #[test]
    fn parse_directory_keeps_v4_stream_size_full_u64() {
        let raw = 0x0000_0001_0000_0100u64; // > 4 GB, legitimate in v4
        let mut data = Vec::new();
        data.extend_from_slice(&make_dir_entry("Root Entry", 5, 0));
        data.extend_from_slice(&make_dir_entry("Big", 2, raw));

        let entries = parse_directory(&data, 4096).expect("parse ok");
        assert_eq!(entries[1].stream_size, raw);
    }

    /// Build a minimal two-entry directory and verify we can locate a stream
    /// regardless of the case of its name. Some third-party Excel-emitting
    /// tools write the workbook stream as `WORKBOOK` instead of `Workbook`.
    #[test]
    fn find_child_by_name_is_case_insensitive() {
        let cfb = CompoundFile {
            file_data: Vec::new(),
            sector_size: 512,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: Vec::new(),
            mini_fat: Vec::new(),
            directory: vec![
                DirectoryEntry {
                    name: "Root Entry".into(),
                    object_type: 5,
                    color: COLOR_BLACK,
                    left_sibling: NOSTREAM,
                    right_sibling: NOSTREAM,
                    child: 1,
                    start_sector: 0,
                    stream_size: 0,
                },
                DirectoryEntry {
                    name: "WORKBOOK".into(),
                    object_type: 2,
                    color: COLOR_BLACK,
                    left_sibling: NOSTREAM,
                    right_sibling: NOSTREAM,
                    child: NOSTREAM,
                    start_sector: 0,
                    stream_size: 0,
                },
            ],
            mini_stream_data: Vec::new(),
        };

        assert!(
            cfb.exists("/Workbook"),
            "canonical case must match WORKBOOK"
        );
        assert!(cfb.exists("/workbook"), "lowercase must match WORKBOOK");
        assert!(cfb.exists("/WORKBOOK"), "exact case must still match");
        assert!(!cfb.exists("/Book"), "unrelated name must not match");
    }
}

const FATSECT: u32 = 0xFFFFFFFD;
const DIFSECT: u32 = 0xFFFFFFFC;
const NUM_DIFAT_ENTRIES_IN_HEADER: usize = 109;
const ROOT_DIR_NAME: &str = "Root Entry";
const MAX_NAME_LEN_UTF16: usize = 31;

const OBJ_TYPE_STORAGE: u8 = 1;
const OBJ_TYPE_STREAM: u8 = 2;
const OBJ_TYPE_ROOT: u8 = 5;
const COLOR_RED: u8 = 0;
const COLOR_BLACK: u8 = 1;

#[derive(Debug)]
enum EntryKind {
    Storage,
    Stream(Vec<u8>),
}

#[derive(Debug)]
struct EntryDef {
    path: String,
    kind: EntryKind,
}

/// One-shot builder for a CFB v3 compound file (512-byte sectors).
///
/// Designed for small, write-once outputs like the OOXML encryption
/// envelope: pile up storages and streams, call [`build`], get bytes.
/// No incremental update / random access — use the existing
/// [`CompoundFile`] reader to read what we wrote.
///
/// Output is byte-compatible with Excel: CFB v3, mini stream cutoff
/// 4096, mini sector size 64. Streams smaller than 4096 bytes go into
/// the mini-FAT; the rest into the regular FAT. Directory entries are
/// emitted as a flat right-sibling chain — readers (Excel, LO, our own
/// reader) walk via DFS so a strict red-black tree is not required.
pub struct CompoundFileBuilder {
    entries: Vec<EntryDef>,
    root_clsid: [u8; 16],
}

impl Default for CompoundFileBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl CompoundFileBuilder {
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
            root_clsid: [0; 16],
        }
    }

    pub fn set_root_clsid(&mut self, clsid: [u8; 16]) {
        self.root_clsid = clsid;
    }

    /// Add an empty storage (directory). Path must start with `/`.
    pub fn add_storage(&mut self, path: &str) -> Result<(), CfbError> {
        validate_path(path)?;
        let segs = split_path(path);
        let name = segs
            .last()
            .ok_or_else(|| CfbError::InvalidFormat(format!("empty path: {path}")))?;
        validate_name(name)?;
        self.entries.push(EntryDef {
            path: path.to_string(),
            kind: EntryKind::Storage,
        });
        Ok(())
    }

    /// Add a stream (file) with the given byte contents. Path must
    /// start with `/`. Stream names follow CFB constraints (≤ 31 UTF-16
    /// code units, no `/`, `\`, `:`, `!`); the `\x06` and `\x05` prefix
    /// bytes used by Office system streams are allowed.
    pub fn add_stream(&mut self, path: &str, data: Vec<u8>) -> Result<(), CfbError> {
        validate_path(path)?;
        let segs = split_path(path);
        let name = segs
            .last()
            .ok_or_else(|| CfbError::InvalidFormat(format!("empty path: {path}")))?;
        validate_name(name)?;
        self.entries.push(EntryDef {
            path: path.to_string(),
            kind: EntryKind::Stream(data),
        });
        Ok(())
    }

    /// Serialize to bytes. Returns the full CFB envelope.
    pub fn build(self) -> Result<Vec<u8>, CfbError> {
        const SECTOR: usize = 512;
        const MINI_SECTOR: usize = 64;
        const MINI_CUTOFF: usize = 4096;
        const ENTRIES_PER_DIR_SECTOR: usize = SECTOR / DIR_ENTRY_LEN;
        const FAT_ENTRIES_PER_SECTOR: usize = SECTOR / 4;

        let mut entries = self.entries;

        // Sort by path depth so parents are always processed before
        // children. Stable sort within a depth keeps insertion order.
        entries.sort_by_key(|e| e.path.matches('/').count());

        // Build directory list with the root entry at index 0. The root
        // is always the first directory entry in CFB.
        let mut dirs: Vec<DirWriteEntry> = vec![DirWriteEntry::root(self.root_clsid)];

        let mut path_to_id: std::collections::HashMap<String, u32> =
            std::collections::HashMap::new();
        path_to_id.insert(String::new(), 0);

        let mut children_of: std::collections::HashMap<u32, Vec<u32>> =
            std::collections::HashMap::new();

        for entry in entries {
            let segments = split_path(&entry.path);
            let parent_path = if segments.len() <= 1 {
                String::new()
            } else {
                format!("/{}", segments[..segments.len() - 1].join("/"))
            };
            let parent_path_key = parent_path.trim_start_matches('/').to_string();
            let name = segments
                .last()
                .ok_or_else(|| CfbError::InvalidFormat(format!("empty path: {}", entry.path)))?
                .to_string();

            let parent_id = *path_to_id.get(&parent_path_key).ok_or_else(|| {
                CfbError::InvalidFormat(format!(
                    "parent path '/{parent_path_key}' not found for '{}'",
                    entry.path
                ))
            })?;

            if !matches!(
                dirs[parent_id as usize].object_type,
                OBJ_TYPE_ROOT | OBJ_TYPE_STORAGE
            ) {
                return Err(CfbError::InvalidFormat(format!(
                    "cannot add '{}' under a stream",
                    entry.path
                )));
            }

            validate_name(&name)?;

            let id = u32::try_from(dirs.len())
                .map_err(|_| CfbError::InvalidFormat("more than 2^32 directory entries".into()))?;

            let (object_type, stream_data) = match entry.kind {
                EntryKind::Storage => (OBJ_TYPE_STORAGE, Vec::new()),
                EntryKind::Stream(data) => (OBJ_TYPE_STREAM, data),
            };

            let key = entry.path.trim_start_matches('/').to_string();
            if path_to_id.contains_key(&key) {
                return Err(CfbError::InvalidFormat(format!(
                    "duplicate entry: {}",
                    entry.path
                )));
            }

            dirs.push(DirWriteEntry {
                name,
                object_type,
                color: COLOR_BLACK,
                left_sibling: NOSTREAM,
                right_sibling: NOSTREAM,
                child: NOSTREAM,
                clsid: [0; 16],
                start_sector: 0,
                stream_size: 0,
                stream_data,
            });
            path_to_id.insert(key, id);
            children_of.entry(parent_id).or_default().push(id);
        }

        // Children form a red-black tree per [MS-CFB] §2.6.4. LibreOffice
        // type detection rejects flat right-sibling chains; we insert in
        // CFB-name order and fix up colors after each insert.
        for (parent_id, child_ids) in children_of.iter_mut() {
            child_ids.sort_by(|&a, &b| {
                compare_cfb_names(&dirs[a as usize].name, &dirs[b as usize].name)
            });
            let root = build_rb_tree(&mut dirs, child_ids);
            dirs[*parent_id as usize].child = root;
        }

        // Partition streams: mini if (kind=stream AND size > 0 AND size < cutoff),
        // regular otherwise. Empty streams keep start_sector=0 stream_size=0.
        let mut mini_stream_data: Vec<u8> = Vec::new();
        let mut mini_chain: Vec<u32> = Vec::new();

        let mut large_streams: Vec<u32> = Vec::new();

        for (idx, dir) in dirs.iter_mut().enumerate() {
            if dir.object_type != OBJ_TYPE_STREAM {
                continue;
            }
            let size = dir.stream_data.len();
            dir.stream_size = size as u64;
            if size == 0 {
                dir.start_sector = 0;
            } else if size < MINI_CUTOFF {
                let first_mini = (mini_stream_data.len() / MINI_SECTOR) as u32;
                dir.start_sector = first_mini;
                let mut written = 0;
                while written < size {
                    let take = (size - written).min(MINI_SECTOR);
                    mini_stream_data.extend_from_slice(&dir.stream_data[written..written + take]);
                    written += take;
                    if written < size {
                        while mini_stream_data.len() % MINI_SECTOR != 0 {
                            mini_stream_data.push(0);
                        }
                        let next = (mini_stream_data.len() / MINI_SECTOR) as u32;
                        mini_chain.push(next);
                    } else {
                        mini_chain.push(ENDOFCHAIN);
                    }
                }
                while mini_stream_data.len() % MINI_SECTOR != 0 {
                    mini_stream_data.push(0);
                }
            } else {
                large_streams.push(idx as u32);
            }
        }

        // Compute sector counts.
        let dir_sector_count = dirs.len().div_ceil(ENTRIES_PER_DIR_SECTOR);
        let mini_fat_sector_count = if mini_chain.is_empty() {
            0
        } else {
            mini_chain.len().div_ceil(FAT_ENTRIES_PER_SECTOR)
        };
        let mini_stream_sector_count = mini_stream_data.len().div_ceil(SECTOR);

        let mut large_sector_counts: Vec<usize> = Vec::with_capacity(large_streams.len());
        for &id in &large_streams {
            let n = (dirs[id as usize].stream_data.len()).div_ceil(SECTOR);
            large_sector_counts.push(n);
        }

        let non_fat_sectors = dir_sector_count
            + mini_fat_sector_count
            + mini_stream_sector_count
            + large_sector_counts.iter().sum::<usize>();

        // FAT and DIFAT counts depend on each other (DIFAT sectors are
        // themselves regular sectors covered by the FAT, but FAT
        // sectors > 109 require DIFAT entries to address). Iterate to a
        // fixpoint — converges in ≤ 2 iterations for any practical file.
        let difat_entries_per_sector = FAT_ENTRIES_PER_SECTOR - 1;
        let mut fat_sector_count = non_fat_sectors.div_ceil(FAT_ENTRIES_PER_SECTOR - 1).max(1);
        let mut difat_sector_count;
        loop {
            difat_sector_count = fat_sector_count
                .saturating_sub(NUM_DIFAT_ENTRIES_IN_HEADER)
                .div_ceil(difat_entries_per_sector);
            let new_fat = (non_fat_sectors + difat_sector_count)
                .div_ceil(FAT_ENTRIES_PER_SECTOR - 1)
                .max(1);
            if new_fat == fat_sector_count {
                break;
            }
            fat_sector_count = new_fat;
        }

        let total_sectors = non_fat_sectors + fat_sector_count + difat_sector_count;

        let mut next_sector: u32 = 0;
        let alloc = |count: usize, next_sector: &mut u32| -> std::ops::Range<u32> {
            let start = *next_sector;
            *next_sector += count as u32;
            start..*next_sector
        };

        let fat_range = alloc(fat_sector_count, &mut next_sector);
        let difat_range = alloc(difat_sector_count, &mut next_sector);
        let dir_range = alloc(dir_sector_count, &mut next_sector);
        let mini_fat_range = alloc(mini_fat_sector_count, &mut next_sector);
        let mini_stream_range = alloc(mini_stream_sector_count, &mut next_sector);
        let mut large_ranges: Vec<std::ops::Range<u32>> = Vec::new();
        for &count in &large_sector_counts {
            large_ranges.push(alloc(count, &mut next_sector));
        }

        let total_fat_entries = fat_sector_count * FAT_ENTRIES_PER_SECTOR;
        let mut fat = vec![FREESECT; total_fat_entries];

        for s in fat_range.clone() {
            fat[s as usize] = FATSECT;
        }
        for s in difat_range.clone() {
            fat[s as usize] = DIFSECT;
        }
        link_chain(&mut fat, dir_range.clone());
        if mini_fat_sector_count > 0 {
            link_chain(&mut fat, mini_fat_range.clone());
        }
        if mini_stream_sector_count > 0 {
            link_chain(&mut fat, mini_stream_range.clone());
        }
        for r in &large_ranges {
            link_chain(&mut fat, r.clone());
        }

        // Wire start_sector for each large stream now that we know its
        // location. Mini streams already had start_sector set (mini
        // sector index, not regular).
        for (i, &id) in large_streams.iter().enumerate() {
            dirs[id as usize].start_sector = large_ranges[i].start;
        }

        // Root entry's start_sector points to the mini stream; its
        // stream_size is the byte length of the mini stream.
        if mini_stream_sector_count > 0 {
            dirs[0].start_sector = mini_stream_range.start;
            dirs[0].stream_size = mini_stream_data.len() as u64;
        }

        // Header.
        let mut buf: Vec<u8> = vec![0u8; (1 + total_sectors as usize) * SECTOR];
        write_header(
            &mut buf,
            fat_range.clone(),
            dir_range.clone(),
            mini_fat_range.clone(),
            difat_range.clone(),
            mini_fat_sector_count as u32,
            dir_sector_count as u32,
            difat_sector_count as u32,
        );

        // FAT sectors.
        {
            let off = (1 + fat_range.start as usize) * SECTOR;
            let mut cursor = off;
            for entry in &fat {
                buf[cursor..cursor + 4].copy_from_slice(&entry.to_le_bytes());
                cursor += 4;
            }
        }

        // DIFAT sectors. Each holds (entries_per_sector - 1) FAT
        // pointers and a "next DIFAT sector" pointer at the tail.
        if difat_sector_count > 0 {
            let header_difat_count = NUM_DIFAT_ENTRIES_IN_HEADER.min(fat_sector_count);
            let mut fat_idx = header_difat_count;
            let difat_sectors: Vec<u32> = difat_range.clone().collect();
            for (i, &sector) in difat_sectors.iter().enumerate() {
                let off = (1 + sector as usize) * SECTOR;
                for slot in 0..difat_entries_per_sector {
                    let val = if fat_idx < fat_sector_count {
                        let v = fat_range.start + fat_idx as u32;
                        fat_idx += 1;
                        v
                    } else {
                        FREESECT
                    };
                    buf[off + slot * 4..off + slot * 4 + 4].copy_from_slice(&val.to_le_bytes());
                }
                let next = if i + 1 < difat_sectors.len() {
                    difat_sectors[i + 1]
                } else {
                    ENDOFCHAIN
                };
                let tail = off + difat_entries_per_sector * 4;
                buf[tail..tail + 4].copy_from_slice(&next.to_le_bytes());
            }
        }

        // Directory sectors. Each dir entry = 128 bytes.
        let dir_byte_offset = (1 + dir_range.start as usize) * SECTOR;
        for (i, dir) in dirs.iter().enumerate() {
            let off = dir_byte_offset + i * DIR_ENTRY_LEN;
            write_dir_entry(&mut buf[off..off + DIR_ENTRY_LEN], dir);
        }
        // Pad remaining directory slots with unallocated entries (zeroed
        // names, object_type=0). The Vec is zero-initialised so they
        // are already correct, but we set color/left/right/child to
        // NOSTREAM per spec for unused entries.
        let used_entries = dirs.len();
        let total_dir_entries = dir_sector_count * ENTRIES_PER_DIR_SECTOR;
        for i in used_entries..total_dir_entries {
            let off = dir_byte_offset + i * DIR_ENTRY_LEN;
            buf[off + 0x44..off + 0x48].copy_from_slice(&NOSTREAM.to_le_bytes());
            buf[off + 0x48..off + 0x4C].copy_from_slice(&NOSTREAM.to_le_bytes());
            buf[off + 0x4C..off + 0x50].copy_from_slice(&NOSTREAM.to_le_bytes());
        }

        // Mini-FAT sectors.
        if mini_fat_sector_count > 0 {
            let off = (1 + mini_fat_range.start as usize) * SECTOR;
            let mut cur = off;
            for entry in &mini_chain {
                buf[cur..cur + 4].copy_from_slice(&entry.to_le_bytes());
                cur += 4;
            }
            // Pad rest of mini-FAT sector(s) with FREESECT.
            let end = (1 + mini_fat_range.end as usize) * SECTOR;
            while cur < end {
                buf[cur..cur + 4].copy_from_slice(&FREESECT.to_le_bytes());
                cur += 4;
            }
        }

        // Mini-stream payload.
        if mini_stream_sector_count > 0 {
            let off = (1 + mini_stream_range.start as usize) * SECTOR;
            buf[off..off + mini_stream_data.len()].copy_from_slice(&mini_stream_data);
        }

        // Large streams.
        for (i, &id) in large_streams.iter().enumerate() {
            let r = &large_ranges[i];
            let off = (1 + r.start as usize) * SECTOR;
            let data = &dirs[id as usize].stream_data;
            buf[off..off + data.len()].copy_from_slice(data);
        }

        Ok(buf)
    }
}

#[derive(Debug)]
struct DirWriteEntry {
    name: String,
    object_type: u8,
    color: u8,
    left_sibling: u32,
    right_sibling: u32,
    child: u32,
    clsid: [u8; 16],
    start_sector: u32,
    stream_size: u64,
    stream_data: Vec<u8>,
}

impl DirWriteEntry {
    fn root(clsid: [u8; 16]) -> Self {
        Self {
            name: ROOT_DIR_NAME.to_string(),
            object_type: OBJ_TYPE_ROOT,
            color: COLOR_BLACK,
            left_sibling: NOSTREAM,
            right_sibling: NOSTREAM,
            child: NOSTREAM,
            clsid,
            start_sector: ENDOFCHAIN,
            stream_size: 0,
            stream_data: Vec::new(),
        }
    }
}

fn validate_path(path: &str) -> Result<(), CfbError> {
    if !path.starts_with('/') {
        return Err(CfbError::InvalidFormat(format!(
            "path must start with '/': {path}"
        )));
    }
    if path == "/" {
        return Err(CfbError::InvalidFormat(
            "cannot add the root entry; it is implicit".into(),
        ));
    }
    Ok(())
}

fn validate_name(name: &str) -> Result<(), CfbError> {
    let utf16_units = name.encode_utf16().count();
    if utf16_units == 0 {
        return Err(CfbError::InvalidFormat("empty name".into()));
    }
    if utf16_units > MAX_NAME_LEN_UTF16 {
        return Err(CfbError::InvalidFormat(format!(
            "name too long ({utf16_units} UTF-16 units, max {MAX_NAME_LEN_UTF16}): {name}"
        )));
    }
    for ch in &['/', '\\', ':', '!'] {
        if name.contains(*ch) {
            return Err(CfbError::InvalidFormat(format!(
                "name cannot contain '{ch}': {name}"
            )));
        }
    }
    Ok(())
}

fn split_path(path: &str) -> Vec<&str> {
    path.trim_start_matches('/')
        .split('/')
        .filter(|s| !s.is_empty())
        .collect()
}

/// CFB name comparison: shorter names first; equal length compared
/// case-insensitively in UTF-16 order. Mirrors MS-CFB §2.6.4. ASCII
/// fast path mirrors the upstream `cfb` crate's optimisation.
fn compare_cfb_names(a: &str, b: &str) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    if a.is_ascii() && b.is_ascii() {
        match a.len().cmp(&b.len()) {
            Ordering::Equal => a
                .bytes()
                .map(|b| b.to_ascii_uppercase())
                .cmp(b.bytes().map(|b| b.to_ascii_uppercase())),
            other => other,
        }
    } else {
        match a.encode_utf16().count().cmp(&b.encode_utf16().count()) {
            Ordering::Equal => a
                .chars()
                .map(|c| c.to_uppercase().next().unwrap_or(c))
                .cmp(b.chars().map(|c| c.to_uppercase().next().unwrap_or(c))),
            other => other,
        }
    }
}

// Standard CLRS §13.3 red-black insert: color new node red, walk parent
// stack performing recolors and rotations until invariants restored.
fn build_rb_tree(dirs: &mut [DirWriteEntry], sorted_ids: &[u32]) -> u32 {
    let mut root = NOSTREAM;
    for &id in sorted_ids {
        dirs[id as usize].color = COLOR_RED;
        dirs[id as usize].left_sibling = NOSTREAM;
        dirs[id as usize].right_sibling = NOSTREAM;
        root = rb_insert(dirs, root, id);
    }
    if root != NOSTREAM {
        dirs[root as usize].color = COLOR_BLACK;
    }
    root
}

fn rb_insert(dirs: &mut [DirWriteEntry], root: u32, new_id: u32) -> u32 {
    let mut parents: Vec<u32> = Vec::with_capacity(16);
    let mut cur = root;
    if cur == NOSTREAM {
        return new_id;
    }
    loop {
        parents.push(cur);
        let cmp = compare_cfb_names(
            &dirs[new_id as usize].name.clone(),
            &dirs[cur as usize].name.clone(),
        );
        let next = match cmp {
            std::cmp::Ordering::Less | std::cmp::Ordering::Equal => dirs[cur as usize].left_sibling,
            std::cmp::Ordering::Greater => dirs[cur as usize].right_sibling,
        };
        if next == NOSTREAM {
            match cmp {
                std::cmp::Ordering::Less | std::cmp::Ordering::Equal => {
                    dirs[cur as usize].left_sibling = new_id;
                }
                std::cmp::Ordering::Greater => {
                    dirs[cur as usize].right_sibling = new_id;
                }
            }
            break;
        }
        cur = next;
    }

    rb_fixup(dirs, root, new_id, &parents)
}

fn rb_fixup(dirs: &mut [DirWriteEntry], mut root: u32, node: u32, parents: &[u32]) -> u32 {
    let mut cur = node;
    let mut stack: Vec<u32> = parents.to_vec();
    while let Some(&parent) = stack.last() {
        if dirs[parent as usize].color == COLOR_BLACK {
            break;
        }
        let grandparent = match stack.iter().rev().nth(1) {
            Some(&g) => g,
            None => break,
        };
        let parent_is_left = dirs[grandparent as usize].left_sibling == parent;
        let uncle = if parent_is_left {
            dirs[grandparent as usize].right_sibling
        } else {
            dirs[grandparent as usize].left_sibling
        };

        if uncle != NOSTREAM && dirs[uncle as usize].color == COLOR_RED {
            dirs[parent as usize].color = COLOR_BLACK;
            dirs[uncle as usize].color = COLOR_BLACK;
            dirs[grandparent as usize].color = COLOR_RED;
            cur = grandparent;
            stack.pop();
            stack.pop();
            continue;
        }

        if parent_is_left {
            if dirs[parent as usize].right_sibling == cur {
                let new_subroot = rotate_left(dirs, parent);
                update_parent_link(dirs, &stack, parent, new_subroot, &mut root, 1);
                stack.pop();
                stack.push(new_subroot);
            }
            let parent2 = *stack.last().unwrap();
            dirs[parent2 as usize].color = COLOR_BLACK;
            dirs[grandparent as usize].color = COLOR_RED;
            let new_subroot = rotate_right(dirs, grandparent);
            update_parent_link(dirs, &stack, grandparent, new_subroot, &mut root, 2);
        } else {
            if dirs[parent as usize].left_sibling == cur {
                let new_subroot = rotate_right(dirs, parent);
                update_parent_link(dirs, &stack, parent, new_subroot, &mut root, 1);
                stack.pop();
                stack.push(new_subroot);
            }
            let parent2 = *stack.last().unwrap();
            dirs[parent2 as usize].color = COLOR_BLACK;
            dirs[grandparent as usize].color = COLOR_RED;
            let new_subroot = rotate_left(dirs, grandparent);
            update_parent_link(dirs, &stack, grandparent, new_subroot, &mut root, 2);
        }
        break;
    }
    root
}

fn rotate_left(dirs: &mut [DirWriteEntry], x: u32) -> u32 {
    let y = dirs[x as usize].right_sibling;
    dirs[x as usize].right_sibling = dirs[y as usize].left_sibling;
    dirs[y as usize].left_sibling = x;
    y
}

fn rotate_right(dirs: &mut [DirWriteEntry], x: u32) -> u32 {
    let y = dirs[x as usize].left_sibling;
    dirs[x as usize].left_sibling = dirs[y as usize].right_sibling;
    dirs[y as usize].right_sibling = x;
    y
}

// `depth_above`: 1 = parent, 2 = grandparent.
fn update_parent_link(
    dirs: &mut [DirWriteEntry],
    stack: &[u32],
    old_subroot: u32,
    new_subroot: u32,
    root: &mut u32,
    depth_above: usize,
) {
    if stack.len() <= depth_above {
        *root = new_subroot;
        return;
    }
    let p = stack[stack.len() - 1 - depth_above];
    if dirs[p as usize].left_sibling == old_subroot {
        dirs[p as usize].left_sibling = new_subroot;
    } else {
        dirs[p as usize].right_sibling = new_subroot;
    }
}

fn link_chain(fat: &mut [u32], range: std::ops::Range<u32>) {
    let count = (range.end - range.start) as usize;
    for i in 0..count {
        let idx = (range.start + i as u32) as usize;
        fat[idx] = if i + 1 == count {
            ENDOFCHAIN
        } else {
            range.start + i as u32 + 1
        };
    }
}

fn write_header(
    buf: &mut [u8],
    fat_range: std::ops::Range<u32>,
    dir_range: std::ops::Range<u32>,
    mini_fat_range: std::ops::Range<u32>,
    difat_range: std::ops::Range<u32>,
    mini_fat_sector_count: u32,
    _dir_sector_count: u32,
    difat_sector_count: u32,
) {
    let fat_sector_count = fat_range.end - fat_range.start;
    // Field offsets per [MS-CFB] §2.2.
    buf[0..8].copy_from_slice(&CFB_MAGIC);
    buf[24..26].copy_from_slice(&0x003Eu16.to_le_bytes());
    buf[26..28].copy_from_slice(&0x0003u16.to_le_bytes());
    buf[28..30].copy_from_slice(&0xFFFEu16.to_le_bytes());
    buf[30..32].copy_from_slice(&9u16.to_le_bytes());
    buf[32..34].copy_from_slice(&6u16.to_le_bytes());
    // [MS-CFB] §2.2: num-directory-sectors MUST be zero in v3.
    // LibreOffice loadenv type detection rejects non-zero values here.
    buf[40..44].copy_from_slice(&0u32.to_le_bytes());
    buf[44..48].copy_from_slice(&fat_sector_count.to_le_bytes());
    buf[48..52].copy_from_slice(&dir_range.start.to_le_bytes());
    buf[52..56].copy_from_slice(&0u32.to_le_bytes());
    buf[56..60].copy_from_slice(&0x0000_1000u32.to_le_bytes());
    let first_mini_fat = if mini_fat_sector_count > 0 {
        mini_fat_range.start
    } else {
        ENDOFCHAIN
    };
    buf[60..64].copy_from_slice(&first_mini_fat.to_le_bytes());
    buf[64..68].copy_from_slice(&mini_fat_sector_count.to_le_bytes());
    let first_difat = if difat_sector_count > 0 {
        difat_range.start
    } else {
        ENDOFCHAIN
    };
    buf[68..72].copy_from_slice(&first_difat.to_le_bytes());
    buf[72..76].copy_from_slice(&difat_sector_count.to_le_bytes());

    let header_difat_count = NUM_DIFAT_ENTRIES_IN_HEADER.min(fat_sector_count as usize);
    for i in 0..NUM_DIFAT_ENTRIES_IN_HEADER {
        let off = 76 + i * 4;
        let val = if i < header_difat_count {
            fat_range.start + i as u32
        } else {
            FREESECT
        };
        buf[off..off + 4].copy_from_slice(&val.to_le_bytes());
    }
}

fn write_dir_entry(buf: &mut [u8], dir: &DirWriteEntry) {
    let utf16: Vec<u16> = dir.name.encode_utf16().collect();
    let name_bytes = utf16.len() * 2;
    for (i, u) in utf16.iter().enumerate() {
        buf[i * 2..i * 2 + 2].copy_from_slice(&u.to_le_bytes());
    }
    buf[name_bytes..name_bytes + 2].copy_from_slice(&0u16.to_le_bytes());
    buf[64..66].copy_from_slice(&((name_bytes + 2) as u16).to_le_bytes());
    buf[66] = dir.object_type;
    buf[67] = dir.color;
    buf[68..72].copy_from_slice(&dir.left_sibling.to_le_bytes());
    buf[72..76].copy_from_slice(&dir.right_sibling.to_le_bytes());
    buf[76..80].copy_from_slice(&dir.child.to_le_bytes());
    buf[80..96].copy_from_slice(&dir.clsid);
    buf[96..100].copy_from_slice(&0u32.to_le_bytes());
    buf[100..108].copy_from_slice(&0u64.to_le_bytes());
    buf[108..116].copy_from_slice(&0u64.to_le_bytes());
    buf[116..120].copy_from_slice(&dir.start_sector.to_le_bytes());
    buf[120..128].copy_from_slice(&dir.stream_size.to_le_bytes());
}

#[cfg(test)]
mod writer_tests {
    use super::*;
    use std::io::Cursor;

    fn build_minimal() -> Vec<u8> {
        let mut b = CompoundFileBuilder::new();
        b.add_stream("/Hello", b"world".to_vec()).unwrap();
        b.build().unwrap()
    }

    #[test]
    fn writer_emits_cfb_v3_magic_and_header() {
        let bytes = build_minimal();
        assert_eq!(&bytes[0..8], &CFB_MAGIC);
        assert_eq!(
            u16::from_le_bytes([bytes[26], bytes[27]]),
            3,
            "major version must be 3"
        );
        assert_eq!(
            u16::from_le_bytes([bytes[30], bytes[31]]),
            9,
            "sector shift must be 9 (512-byte sectors)"
        );
    }

    #[test]
    fn writer_round_trips_through_reader() {
        let bytes = build_minimal();
        let cfb = CompoundFile::open(Cursor::new(&bytes))
            .expect("our reader must accept our writer's output");
        assert!(cfb.exists("/Hello"));
        assert_eq!(cfb.read_stream("/Hello").unwrap(), b"world");
    }

    #[test]
    fn writer_emits_configured_root_clsid() {
        let clsid = [
            0x20, 0x08, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x46,
        ];
        let mut builder = CompoundFileBuilder::new();
        builder.set_root_clsid(clsid);
        builder.add_stream("/Workbook", b"data".to_vec()).unwrap();

        let bytes = builder.build().unwrap();
        let first_dir_sector =
            u32::from_le_bytes([bytes[48], bytes[49], bytes[50], bytes[51]]) as usize;
        let root_entry = (1 + first_dir_sector) * 512;

        assert_eq!(&bytes[root_entry + 80..root_entry + 96], &clsid);
    }

    #[test]
    fn writer_handles_storages_and_nested_streams() {
        let mut b = CompoundFileBuilder::new();
        b.add_storage("/Outer").unwrap();
        b.add_storage("/Outer/Inner").unwrap();
        b.add_stream("/Outer/Inner/leaf.txt", b"deeply nested".to_vec())
            .unwrap();
        b.add_stream("/top.bin", vec![0xAA; 100]).unwrap();
        let bytes = b.build().unwrap();
        let cfb = CompoundFile::open(Cursor::new(&bytes)).unwrap();
        // `exists` only matches streams; storages aren't user-visible
        // through that API. Verify them through directory inspection.
        let names: Vec<_> = cfb
            .directory
            .iter()
            .map(|e| (e.name.clone(), e.object_type))
            .collect();
        assert!(names.contains(&("Outer".to_string(), OBJ_TYPE_STORAGE)));
        assert!(names.contains(&("Inner".to_string(), OBJ_TYPE_STORAGE)));
        assert!(cfb.exists("/Outer/Inner/leaf.txt"));
        assert!(cfb.exists("/top.bin"));
        assert_eq!(
            cfb.read_stream("/Outer/Inner/leaf.txt").unwrap(),
            b"deeply nested"
        );
        assert_eq!(cfb.read_stream("/top.bin").unwrap(), vec![0xAA; 100]);
    }

    #[test]
    fn writer_routes_small_streams_to_mini_fat_and_large_to_regular() {
        let mut b = CompoundFileBuilder::new();
        let small = vec![0x42; 1000];
        let large = vec![0x99; 5000];
        b.add_stream("/small", small.clone()).unwrap();
        b.add_stream("/large", large.clone()).unwrap();
        let bytes = b.build().unwrap();
        let cfb = CompoundFile::open(Cursor::new(&bytes)).unwrap();
        assert_eq!(cfb.read_stream("/small").unwrap(), small);
        assert_eq!(cfb.read_stream("/large").unwrap(), large);
    }

    #[test]
    fn writer_handles_streams_at_mini_cutoff_boundary() {
        // A 4096-byte stream goes to regular FAT (not mini), per the
        // MS-CFB rule "stream size >= cutoff goes to regular FAT".
        let mut b = CompoundFileBuilder::new();
        b.add_stream("/at_cutoff", vec![1u8; 4096]).unwrap();
        b.add_stream("/just_under", vec![2u8; 4095]).unwrap();
        let bytes = b.build().unwrap();
        let cfb = CompoundFile::open(Cursor::new(&bytes)).unwrap();
        assert_eq!(cfb.read_stream("/at_cutoff").unwrap(), vec![1u8; 4096]);
        assert_eq!(cfb.read_stream("/just_under").unwrap(), vec![2u8; 4095]);
    }

    #[test]
    fn writer_supports_control_character_prefix_in_name() {
        let mut b = CompoundFileBuilder::new();
        b.add_storage("/\u{0006}DataSpaces").unwrap();
        b.add_stream("/\u{0006}DataSpaces/Version", b"v1.0".to_vec())
            .unwrap();
        let bytes = b.build().unwrap();
        let cfb = CompoundFile::open(Cursor::new(&bytes)).unwrap();
        // The reader's `exists` API only matches streams, not
        // storages — so verify the storage exists by inspecting the
        // directory directly, and verify the nested stream via
        // `exists` + `read_stream`.
        assert!(
            cfb.directory
                .iter()
                .any(|e| e.name == "\u{0006}DataSpaces" && e.object_type == OBJ_TYPE_STORAGE),
            "storage with control-char prefix must be present in directory"
        );
        assert!(cfb.exists("/\u{0006}DataSpaces/Version"));
        assert_eq!(
            cfb.read_stream("/\u{0006}DataSpaces/Version").unwrap(),
            b"v1.0"
        );
    }

    #[test]
    fn writer_rejects_duplicates() {
        let mut b = CompoundFileBuilder::new();
        b.add_stream("/dup", b"a".to_vec()).unwrap();
        b.add_stream("/dup", b"b".to_vec()).unwrap();
        let err = b.build().expect_err("duplicates must error");
        assert!(matches!(err, CfbError::InvalidFormat(msg) if msg.contains("duplicate")));
    }

    #[test]
    fn writer_rejects_long_names() {
        let mut b = CompoundFileBuilder::new();
        let long = "a".repeat(32);
        let err = b.add_stream(&format!("/{long}"), vec![]).unwrap_err();
        assert!(matches!(err, CfbError::InvalidFormat(_)));
    }

    #[test]
    fn writer_rejects_streams_under_streams() {
        let mut b = CompoundFileBuilder::new();
        b.add_stream("/leaf", b"x".to_vec()).unwrap();
        b.add_stream("/leaf/below", b"y".to_vec()).unwrap();
        let err = b.build().expect_err("can't nest under a stream");
        assert!(matches!(err, CfbError::InvalidFormat(_)));
    }

    #[test]
    fn writer_handles_difat_chain_for_large_streams() {
        // 8 MB stream forces > 109 FAT sectors. Each FAT sector covers
        // 128 sectors × 512 B = 64 KiB; 8 MiB / 64 KiB = 128 sectors,
        // so we need at least 1 DIFAT sector.
        const SIZE: usize = 8 * 1024 * 1024;
        let mut b = CompoundFileBuilder::new();
        let payload: Vec<u8> = (0..SIZE).map(|i| (i & 0xFF) as u8).collect();
        b.add_stream("/big", payload.clone()).unwrap();
        let bytes = b.build().unwrap();

        let num_fat = u32::from_le_bytes([bytes[44], bytes[45], bytes[46], bytes[47]]);
        let num_difat = u32::from_le_bytes([bytes[72], bytes[73], bytes[74], bytes[75]]);
        let first_difat = u32::from_le_bytes([bytes[68], bytes[69], bytes[70], bytes[71]]);
        assert!(num_fat > NUM_DIFAT_ENTRIES_IN_HEADER as u32);
        assert!(num_difat >= 1);
        assert_ne!(first_difat, ENDOFCHAIN);

        let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("reader must accept DIFAT chain");
        assert_eq!(cfb.read_stream("/big").unwrap(), payload);
    }

    #[test]
    fn writer_rb_tree_round_trips_50_children_with_invariants_intact() {
        let mut b = CompoundFileBuilder::new();
        for i in 0..50u32 {
            b.add_stream(&format!("/s{i:03}"), vec![i as u8; 8])
                .unwrap();
        }
        let bytes = b.build().unwrap();
        let cfb = CompoundFile::open(Cursor::new(&bytes)).unwrap();
        for i in 0..50u32 {
            let path = format!("/s{i:03}");
            assert!(cfb.exists(&path), "missing {path}");
            assert_eq!(cfb.read_stream(&path).unwrap(), vec![i as u8; 8]);
        }
        assert_rb_invariants(&cfb);
    }

    #[test]
    fn writer_rb_tree_holds_invariants_for_fan_outs_three_through_seven() {
        for n in 3..=7u32 {
            let mut b = CompoundFileBuilder::new();
            b.add_storage("/grp").unwrap();
            for i in 0..n {
                b.add_stream(&format!("/grp/s{i}"), vec![]).unwrap();
            }
            let bytes = b.build().unwrap();
            let cfb =
                CompoundFile::open(Cursor::new(&bytes)).unwrap_or_else(|e| panic!("n={n}: {e}"));
            assert_rb_invariants(&cfb);
        }
    }

    fn assert_rb_invariants(cfb: &CompoundFile) {
        let entries: Vec<&DirectoryEntry> = cfb.directory_entries().collect();
        for entry in &entries {
            if !matches!(entry.object_type, OBJ_TYPE_ROOT | OBJ_TYPE_STORAGE) {
                continue;
            }
            let root = entry.child;
            if root == NOSTREAM {
                continue;
            }
            assert_eq!(
                entries[root as usize].color, COLOR_BLACK,
                "subtree root under {:?} must be black",
                entry.name
            );
            assert_rb_subtree(&entries, root);
        }
    }

    // CLRS §13.1: red parent → black children; equal black-height on
    // every root-to-NIL path. Returns black-height for caller's check.
    fn assert_rb_subtree(entries: &[&DirectoryEntry], node: u32) -> u32 {
        if node == NOSTREAM {
            return 1;
        }
        let e = entries[node as usize];
        if e.color == COLOR_RED {
            for child in [e.left_sibling, e.right_sibling] {
                if child != NOSTREAM {
                    assert_eq!(
                        entries[child as usize].color, COLOR_BLACK,
                        "red node {} has red child {}",
                        e.name, entries[child as usize].name
                    );
                }
            }
        }
        let lh = assert_rb_subtree(entries, e.left_sibling);
        let rh = assert_rb_subtree(entries, e.right_sibling);
        assert_eq!(
            lh, rh,
            "black-height mismatch at {}: left={lh} right={rh}",
            e.name
        );
        lh + u32::from(e.color == COLOR_BLACK)
    }
}
