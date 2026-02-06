//! CSV writer

use std::fs::File;
use std::io::Write;
use std::path::Path;

use crate::error::{CsvError, CsvResult};
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
            LineTerminator::LF => csv::Terminator::Any(b'\n'),
            LineTerminator::CRLF => csv::Terminator::CRLF,
            LineTerminator::CR => csv::Terminator::Any(b'\r'),
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
}
