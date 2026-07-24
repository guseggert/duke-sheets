//! CSV writer

use std::fs::File;
use std::io::Write;
use std::path::Path;

use crate::error::CsvResult;
use crate::options::{CsvWriteOptions, LineTerminator};
use duke_sheets_core::Worksheet;

/// CSV file writer
pub struct CsvWriter;

impl CsvWriter {
    /// Write a worksheet to a CSV file
    pub fn write_file<P: AsRef<Path>>(
        worksheet: &Worksheet,
        path: P,
        options: &CsvWriteOptions,
    ) -> CsvResult<()> {
        let file = File::create(path)?;
        Self::write(worksheet, file, options)
    }

    /// Write a worksheet to a writer
    pub fn write<W: Write>(
        worksheet: &Worksheet,
        writer: W,
        options: &CsvWriteOptions,
    ) -> CsvResult<()> {
        let terminator = match options.line_terminator {
            LineTerminator::Lf => csv::Terminator::Any(b'\n'),
            LineTerminator::Crlf => csv::Terminator::CRLF,
            LineTerminator::Cr => csv::Terminator::Any(b'\r'),
        };

        let mut csv_writer = csv::WriterBuilder::new()
            .delimiter(options.delimiter)
            .quote(options.quote)
            .terminator(terminator)
            .from_writer(writer);

        if let Some(range) = worksheet.used_range() {
            for row in range.start.row..=range.end.row {
                let mut record = Vec::new();

                for col in range.start.col..=range.end.col {
                    let value = worksheet.get_value_at(row, col);
                    record.push(value.to_string());
                }

                csv_writer.write_record(&record)?;
            }
        }

        csv_writer.flush()?;
        Ok(())
    }

    /// Write a worksheet to a CSV string.
    ///
    /// This uses the same writer path as [`Self::write`] so string and byte
    /// destinations have identical CSV output.
    pub fn write_string(worksheet: &Worksheet, options: &CsvWriteOptions) -> CsvResult<String> {
        let mut bytes = Vec::new();
        Self::write(worksheet, &mut bytes, options)?;
        Ok(String::from_utf8(bytes)?)
    }
}

#[cfg(test)]
mod tests {
    use duke_sheets_core::{CellValue, Worksheet};

    use crate::{CsvError, CsvWriteOptions};

    use super::CsvWriter;

    fn sample_worksheet() -> Worksheet {
        let mut worksheet = Worksheet::new("Sheet1");
        worksheet
            .set_cell_value_at(0, 0, CellValue::string("plain"))
            .unwrap();
        worksheet
            .set_cell_value_at(0, 1, CellValue::string("say \"hi\", 世界"))
            .unwrap();
        worksheet
            .set_cell_value_at(1, 0, CellValue::string("line 1\nline 2"))
            .unwrap();
        worksheet
            .set_cell_value_at(1, 1, CellValue::string("😀"))
            .unwrap();
        worksheet
    }

    #[test]
    fn write_string_matches_write_for_unicode_and_quoted_values() {
        let worksheet = sample_worksheet();
        let options = CsvWriteOptions::default();
        let mut bytes = Vec::new();
        CsvWriter::write(&worksheet, &mut bytes, &options).unwrap();

        assert_eq!(
            CsvWriter::write_string(&worksheet, &options).unwrap(),
            String::from_utf8(bytes).unwrap()
        );
        assert_eq!(
            CsvWriter::write_string(&worksheet, &options).unwrap(),
            "plain,\"say \"\"hi\"\", 世界\"\r\n\"line 1\nline 2\",😀\r\n"
        );
    }

    #[test]
    fn write_string_reports_invalid_utf8() {
        let mut worksheet = Worksheet::new("Sheet1");
        worksheet
            .set_cell_value_at(0, 0, CellValue::string("first"))
            .unwrap();
        worksheet
            .set_cell_value_at(0, 1, CellValue::string("second"))
            .unwrap();
        let options = CsvWriteOptions {
            delimiter: 0xff,
            ..CsvWriteOptions::default()
        };

        assert!(matches!(
            CsvWriter::write_string(&worksheet, &options),
            Err(CsvError::Utf8(_))
        ));
    }
}
