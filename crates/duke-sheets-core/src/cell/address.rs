//! Cell address and range types

use crate::error::{Error, Result};
use crate::{MAX_COLS, MAX_ROWS};
use std::fmt;
use std::str::FromStr;

/// A cell address (e.g., "A1", "$B$2")
///
/// Cell addresses in Excel use a combination of column letters (A-XFD) and row numbers (1-1048576).
/// The optional `$` prefix makes a reference absolute (doesn't change when copied).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellAddress {
    /// Row index (0-based internally, 1-based in display)
    pub row: u32,
    /// Column index (0-based, A=0, B=1, ..., XFD=16383)
    pub col: u16,
    /// Whether the row reference is absolute ($)
    pub row_absolute: bool,
    /// Whether the column reference is absolute ($)
    pub col_absolute: bool,
}

impl CellAddress {
    /// Create a new cell address with relative references
    pub fn new(row: u32, col: u16) -> Self {
        Self {
            row,
            col,
            row_absolute: false,
            col_absolute: false,
        }
    }

    /// Create a new cell address with specified absolute/relative flags
    pub fn with_absolute(row: u32, col: u16, row_absolute: bool, col_absolute: bool) -> Self {
        Self {
            row,
            col,
            row_absolute,
            col_absolute,
        }
    }

    /// Create an absolute cell address ($A$1 style)
    pub fn absolute(row: u32, col: u16) -> Self {
        Self {
            row,
            col,
            row_absolute: true,
            col_absolute: true,
        }
    }

    /// Parse a cell address from A1-style notation
    ///
    /// # Examples
    /// ```
    /// use duke_sheets_core::CellAddress;
    ///
    /// let addr = CellAddress::parse("A1").unwrap();
    /// assert_eq!(addr.row, 0);
    /// assert_eq!(addr.col, 0);
    ///
    /// let addr = CellAddress::parse("$B$2").unwrap();
    /// assert_eq!(addr.row, 1);
    /// assert_eq!(addr.col, 1);
    /// assert!(addr.row_absolute);
    /// assert!(addr.col_absolute);
    /// ```
    pub fn parse(s: &str) -> Result<Self> {
        let s = s.trim();
        if s.is_empty() {
            return Err(Error::InvalidAddress("empty address".into()));
        }

        let bytes = s.as_bytes();
        let mut pos = 0;

        // Check for column absolute marker
        let col_absolute = if bytes.get(pos) == Some(&b'$') {
            pos += 1;
            true
        } else {
            false
        };

        // Parse column letters
        let col_start = pos;
        while pos < bytes.len() && bytes[pos].is_ascii_alphabetic() {
            pos += 1;
        }

        if pos == col_start {
            return Err(Error::InvalidAddress(format!(
                "no column letters in '{}'",
                s
            )));
        }

        let col_str = &s[col_start..pos];
        let col = Self::letters_to_column(col_str)?;

        // Check for row absolute marker
        let row_absolute = if bytes.get(pos) == Some(&b'$') {
            pos += 1;
            true
        } else {
            false
        };

        // Parse row number
        let row_str = &s[pos..];
        if row_str.is_empty() {
            return Err(Error::InvalidAddress(format!("no row number in '{}'", s)));
        }

        let row: u32 = row_str
            .parse()
            .map_err(|_| Error::InvalidAddress(format!("invalid row number in '{}'", s)))?;

        // Excel rows are 1-based, we use 0-based internally
        if row == 0 {
            return Err(Error::InvalidAddress(format!(
                "row number must be >= 1 in '{}'",
                s
            )));
        }

        let row = row - 1;

        if row >= MAX_ROWS {
            return Err(Error::RowOutOfBounds(row, MAX_ROWS - 1));
        }

        if col >= MAX_COLS {
            return Err(Error::ColumnOutOfBounds(col, MAX_COLS - 1));
        }

        Ok(Self {
            row,
            col,
            row_absolute,
            col_absolute,
        })
    }

    /// Convert column index to letters (0 = A, 25 = Z, 26 = AA, etc.)
    pub fn column_to_letters(col: u16) -> String {
        let mut result = String::new();
        let mut n = col as u32 + 1; // 1-based for calculation

        while n > 0 {
            n -= 1;
            let c = ((n % 26) as u8 + b'A') as char;
            result.insert(0, c);
            n /= 26;
        }

        result
    }

    /// Convert column letters to index (A = 0, Z = 25, AA = 26, etc.)
    pub fn letters_to_column(letters: &str) -> Result<u16> {
        if letters.is_empty() {
            return Err(Error::InvalidAddress("empty column letters".into()));
        }

        let mut col: u32 = 0;
        for c in letters.chars() {
            if !c.is_ascii_alphabetic() {
                return Err(Error::InvalidAddress(format!(
                    "invalid column letter '{}'",
                    c
                )));
            }
            col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        }

        let col = col - 1; // Convert to 0-based

        if col >= MAX_COLS as u32 {
            return Err(Error::ColumnOutOfBounds(col as u16, MAX_COLS - 1));
        }

        Ok(col as u16)
    }

    /// Format as A1-style string
    pub fn to_a1_string(&self) -> String {
        let mut result = String::new();

        if self.col_absolute {
            result.push('$');
        }
        result.push_str(&Self::column_to_letters(self.col));

        if self.row_absolute {
            result.push('$');
        }
        result.push_str(&(self.row + 1).to_string());

        result
    }

    /// Create a range from this address to another
    pub fn to(&self, other: CellAddress) -> CellRange {
        CellRange::new(*self, other)
    }
}

impl fmt::Display for CellAddress {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_a1_string())
    }
}

impl FromStr for CellAddress {
    type Err = Error;

    fn from_str(s: &str) -> Result<Self> {
        Self::parse(s)
    }
}

/// A range of cells (e.g., "A1:B10")
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRange {
    /// Start address (top-left)
    pub start: CellAddress,
    /// End address (bottom-right)
    pub end: CellAddress,
}

impl CellRange {
    /// Create a new cell range
    pub fn new(start: CellAddress, end: CellAddress) -> Self {
        // Normalize so start is top-left and end is bottom-right
        let (start_row, end_row) = if start.row <= end.row {
            (start.row, end.row)
        } else {
            (end.row, start.row)
        };

        let (start_col, end_col) = if start.col <= end.col {
            (start.col, end.col)
        } else {
            (end.col, start.col)
        };

        Self {
            start: CellAddress::with_absolute(
                start_row,
                start_col,
                start.row_absolute,
                start.col_absolute,
            ),
            end: CellAddress::with_absolute(end_row, end_col, end.row_absolute, end.col_absolute),
        }
    }

    /// Create a range from row/column indices
    pub fn from_indices(start_row: u32, start_col: u16, end_row: u32, end_col: u16) -> Self {
        Self::new(
            CellAddress::new(start_row, start_col),
            CellAddress::new(end_row, end_col),
        )
    }

    /// Create a single-cell range
    pub fn single(addr: CellAddress) -> Self {
        Self {
            start: addr,
            end: addr,
        }
    }

    /// Parse a range from A1:B10 notation
    pub fn parse(s: &str) -> Result<Self> {
        let s = s.trim();

        if let Some(colon_pos) = s.find(':') {
            let start = CellAddress::parse(&s[..colon_pos])?;
            let end = CellAddress::parse(&s[colon_pos + 1..])?;
            Ok(Self::new(start, end))
        } else {
            // Single cell range
            let addr = CellAddress::parse(s)?;
            Ok(Self::single(addr))
        }
    }

    /// Check if a cell is within this range
    pub fn contains(&self, addr: &CellAddress) -> bool {
        addr.row >= self.start.row
            && addr.row <= self.end.row
            && addr.col >= self.start.col
            && addr.col <= self.end.col
    }

    /// Get the number of rows in the range
    pub fn row_count(&self) -> u32 {
        self.end.row - self.start.row + 1
    }

    /// Get the number of columns in the range
    pub fn col_count(&self) -> u16 {
        self.end.col - self.start.col + 1
    }

    /// Get the total number of cells in the range
    pub fn cell_count(&self) -> u64 {
        self.row_count() as u64 * self.col_count() as u64
    }

    /// Check if this range overlaps with another
    pub fn overlaps(&self, other: &CellRange) -> bool {
        self.start.row <= other.end.row
            && self.end.row >= other.start.row
            && self.start.col <= other.end.col
            && self.end.col >= other.start.col
    }

    /// Get the intersection of two ranges, if any
    pub fn intersect(&self, other: &CellRange) -> Option<CellRange> {
        if !self.overlaps(other) {
            return None;
        }

        Some(CellRange::from_indices(
            self.start.row.max(other.start.row),
            self.start.col.max(other.start.col),
            self.end.row.min(other.end.row),
            self.end.col.min(other.end.col),
        ))
    }

    /// Iterate over all cell addresses in the range (row by row)
    pub fn cells(&self) -> CellRangeIterator {
        CellRangeIterator {
            range: *self,
            current_row: self.start.row,
            current_col: self.start.col,
        }
    }

    /// Format as A1:B10 string
    pub fn to_a1_string(&self) -> String {
        if self.start == self.end {
            self.start.to_a1_string()
        } else {
            format!("{}:{}", self.start.to_a1_string(), self.end.to_a1_string())
        }
    }
}

impl fmt::Display for CellRange {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_a1_string())
    }
}

impl FromStr for CellRange {
    type Err = Error;

    fn from_str(s: &str) -> Result<Self> {
        Self::parse(s)
    }
}

/// Iterator over cells in a range
pub struct CellRangeIterator {
    range: CellRange,
    current_row: u32,
    current_col: u16,
}

impl Iterator for CellRangeIterator {
    type Item = CellAddress;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_row > self.range.end.row {
            return None;
        }

        let addr = CellAddress::new(self.current_row, self.current_col);

        // Move to next cell
        self.current_col += 1;
        if self.current_col > self.range.end.col {
            self.current_col = self.range.start.col;
            self.current_row += 1;
        }

        Some(addr)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.range.cell_count() as usize;
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for CellRangeIterator {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_to_letters() {
        assert_eq!(CellAddress::column_to_letters(0), "A");
        assert_eq!(CellAddress::column_to_letters(1), "B");
        assert_eq!(CellAddress::column_to_letters(25), "Z");
        assert_eq!(CellAddress::column_to_letters(26), "AA");
        assert_eq!(CellAddress::column_to_letters(27), "AB");
        assert_eq!(CellAddress::column_to_letters(701), "ZZ");
        assert_eq!(CellAddress::column_to_letters(702), "AAA");
        assert_eq!(CellAddress::column_to_letters(16383), "XFD"); // Max Excel column
    }

    #[test]
    fn test_letters_to_column() {
        assert_eq!(CellAddress::letters_to_column("A").unwrap(), 0);
        assert_eq!(CellAddress::letters_to_column("B").unwrap(), 1);
        assert_eq!(CellAddress::letters_to_column("Z").unwrap(), 25);
        assert_eq!(CellAddress::letters_to_column("AA").unwrap(), 26);
        assert_eq!(CellAddress::letters_to_column("AB").unwrap(), 27);
        assert_eq!(CellAddress::letters_to_column("ZZ").unwrap(), 701);
        assert_eq!(CellAddress::letters_to_column("AAA").unwrap(), 702);
        assert_eq!(CellAddress::letters_to_column("XFD").unwrap(), 16383);

        // Case insensitive
        assert_eq!(CellAddress::letters_to_column("a").unwrap(), 0);
        assert_eq!(CellAddress::letters_to_column("aa").unwrap(), 26);
    }

    #[test]
    fn test_cell_address_parse() {
        let addr = CellAddress::parse("A1").unwrap();
        assert_eq!(addr.row, 0);
        assert_eq!(addr.col, 0);
        assert!(!addr.row_absolute);
        assert!(!addr.col_absolute);

        let addr = CellAddress::parse("B2").unwrap();
        assert_eq!(addr.row, 1);
        assert_eq!(addr.col, 1);

        let addr = CellAddress::parse("$A$1").unwrap();
        assert_eq!(addr.row, 0);
        assert_eq!(addr.col, 0);
        assert!(addr.row_absolute);
        assert!(addr.col_absolute);

        let addr = CellAddress::parse("$A1").unwrap();
        assert!(addr.col_absolute);
        assert!(!addr.row_absolute);

        let addr = CellAddress::parse("A$1").unwrap();
        assert!(!addr.col_absolute);
        assert!(addr.row_absolute);

        let addr = CellAddress::parse("XFD1048576").unwrap();
        assert_eq!(addr.row, 1048575);
        assert_eq!(addr.col, 16383);
    }

    #[test]
    fn test_cell_address_parse_errors() {
        assert!(CellAddress::parse("").is_err());
        assert!(CellAddress::parse("A").is_err());
        assert!(CellAddress::parse("1").is_err());
        assert!(CellAddress::parse("A0").is_err()); // Row 0 is invalid
        assert!(CellAddress::parse("A1048577").is_err()); // Row too large
        assert!(CellAddress::parse("XFE1").is_err()); // Column too large
    }

    #[test]
    fn test_cell_address_display() {
        assert_eq!(CellAddress::new(0, 0).to_string(), "A1");
        assert_eq!(CellAddress::new(99, 2).to_string(), "C100");
        assert_eq!(CellAddress::absolute(0, 0).to_string(), "$A$1");
    }

    #[test]
    fn test_cell_range_parse() {
        let range = CellRange::parse("A1:B2").unwrap();
        assert_eq!(range.start, CellAddress::new(0, 0));
        assert_eq!(range.end, CellAddress::new(1, 1));

        // Single cell
        let range = CellRange::parse("C3").unwrap();
        assert_eq!(range.start, CellAddress::new(2, 2));
        assert_eq!(range.end, CellAddress::new(2, 2));
    }

    #[test]
    fn test_cell_range_contains() {
        let range = CellRange::parse("B2:D4").unwrap();

        assert!(range.contains(&CellAddress::new(1, 1))); // B2
        assert!(range.contains(&CellAddress::new(3, 3))); // D4
        assert!(range.contains(&CellAddress::new(2, 2))); // C3

        assert!(!range.contains(&CellAddress::new(0, 0))); // A1
        assert!(!range.contains(&CellAddress::new(4, 1))); // B5
    }

    #[test]
    fn test_cell_range_iterator() {
        let range = CellRange::parse("A1:B2").unwrap();
        let cells: Vec<_> = range.cells().collect();

        assert_eq!(cells.len(), 4);
        assert_eq!(cells[0], CellAddress::new(0, 0)); // A1
        assert_eq!(cells[1], CellAddress::new(0, 1)); // B1
        assert_eq!(cells[2], CellAddress::new(1, 0)); // A2
        assert_eq!(cells[3], CellAddress::new(1, 1)); // B2
    }
}
