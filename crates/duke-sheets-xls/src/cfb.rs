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
struct DirectoryEntry {
    name: String,
    object_type: u8,
    left_sibling: u32,
    right_sibling: u32,
    child: u32,
    start_sector: u32,
    stream_size: u64,
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

        let total_fat_sectors = read_u32(&file_data, 44)? as usize;

        // Sanity check: FAT sectors can't exceed the number of sectors in the file
        let max_sectors = file_data.len() / sector_size;
        if total_fat_sectors > max_sectors {
            return Err(CfbError::InvalidFormat(format!(
                "total FAT sectors ({total_fat_sectors}) exceeds file capacity ({max_sectors} sectors)"
            )));
        }
        let first_directory_sector = read_u32(&file_data, 48)?;
        let mini_stream_cutoff = read_u32(&file_data, 56)? as usize;
        let first_mini_fat_sector = read_u32(&file_data, 60)?;
        let total_mini_fat_sectors = read_u32(&file_data, 64)? as usize;
        let first_difat_sector = read_u32(&file_data, 68)?;
        let total_difat_sectors = read_u32(&file_data, 72)? as usize;

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
        let directory = parse_directory(&directory_stream)?;
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
        let mini_stream_data = if root_size == 0 || root.start_sector == ENDOFCHAIN {
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
            if entry.name == name {
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

fn read_difat(
    file_data: &[u8],
    sector_size: usize,
    total_fat_sectors: usize,
    first_difat_sector: u32,
    total_difat_sectors: usize,
) -> Result<Vec<u32>, CfbError> {
    let mut difat = Vec::with_capacity(total_fat_sectors);

    for i in 0..109usize {
        let offset = 76 + i * 4;
        let sid = read_u32(file_data, offset)?;
        if sid != FREESECT {
            difat.push(sid);
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

fn parse_directory(data: &[u8]) -> Result<Vec<DirectoryEntry>, CfbError> {
    if data.len() < DIR_ENTRY_LEN {
        return Err(CfbError::InvalidFormat("directory stream too short".into()));
    }

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

        entries.push(DirectoryEntry {
            name,
            object_type: chunk[66],
            left_sibling: read_u32(chunk, 68)?,
            right_sibling: read_u32(chunk, 72)?,
            child: read_u32(chunk, 76)?,
            start_sector: read_u32(chunk, 116)?,
            stream_size: read_u64(chunk, 120)?,
        });

        offset += DIR_ENTRY_LEN;
    }

    Ok(entries)
}
