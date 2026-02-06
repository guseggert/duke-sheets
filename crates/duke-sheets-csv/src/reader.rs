//! CSV reader

use std::fs::File;
use std::io::Read;
use std::path::Path;

use crate::error::{CsvError, CsvResult};
use crate::options::CsvReadOptions;
use duke_sheets_core::{CellValue, Worksheet};

/// CSV file reader
pub struct CsvReader;

impl CsvReader {
    /// Read CSV file into a worksheet
    pub fn read_file<P: AsRef<Path>>(path: P, options: &CsvReadOptions) -> CsvResult<Worksheet> {
        let file = File::open(path)?;
        Self::read(file, options)
    }

    /// Read CSV from a reader into a worksheet
    pub fn read<R: Read>(reader: R, options: &CsvReadOptions) -> CsvResult<Worksheet> {
        let mut csv_reader = csv::ReaderBuilder::new()
            .delimiter(options.delimiter)
            .quote(options.quote)
            .has_headers(options.has_header)
            .from_reader(reader);

        let mut worksheet = Worksheet::new("Sheet1");
        let mut row_idx = 0u32;

        // Read headers if present
        if options.has_header {
            if let Some(headers) = csv_reader.headers().ok() {
                for (col, value) in headers.iter().enumerate() {
                    worksheet.set_cell_value_at(row_idx, col as u16, value)?;
                }
                row_idx += 1;
            }
        }

        // Read records
        for result in csv_reader.records() {
            let record = result?;

            for (col, field) in record.iter().enumerate() {
                let value = if options.auto_detect_types {
                    Self::detect_type(field)
                } else {
                    CellValue::string(field)
                };

                worksheet.set_cell_value_at(row_idx, col as u16, value)?;
            }

            row_idx += 1;
        }

        Ok(worksheet)
    }

    /// Detect the type of a field value
    fn detect_type(field: &str) -> CellValue {
        let field = field.trim();

        if field.is_empty() {
            return CellValue::Empty;
        }

        // Try boolean
        match field.to_lowercase().as_str() {
            "true" | "yes" | "1" => return CellValue::Boolean(true),
            "false" | "no" | "0" if !field.chars().all(|c| c.is_ascii_digit()) => {
                return CellValue::Boolean(false)
            }
            _ => {}
        }

        // Try number
        if let Ok(n) = field.parse::<f64>() {
            return CellValue::Number(n);
        }

        // Default to string
        CellValue::string(field)
    }
}
