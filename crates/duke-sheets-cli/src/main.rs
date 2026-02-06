//! Duke Sheets CLI - spreadsheet conversion tool

use anyhow::{Context, Result};
use clap::{Parser, Subcommand};
use duke_sheets::prelude::*;
use duke_sheets::{CalculationOptions, WorkbookCalculationExt};
use std::io::{self, Write};
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "duke")]
#[command(
    author,
    version,
    about = "Spreadsheet conversion and manipulation tool"
)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Convert a spreadsheet to CSV and output to stdout or file
    #[command(alias = "csv")]
    ToCsv {
        /// Input spreadsheet file (xlsx, xls, csv)
        input: PathBuf,

        /// Output CSV file (default: stdout)
        #[arg(short, long)]
        output: Option<PathBuf>,

        /// Sheet index to convert (0-based, default: 0)
        #[arg(short, long, default_value = "0")]
        sheet: usize,

        /// Calculate formulas before export
        #[arg(short, long)]
        calculate: bool,

        /// Field delimiter (default: comma)
        #[arg(short, long, default_value = ",")]
        delimiter: char,
    },

    /// Show information about a spreadsheet
    Info {
        /// Input spreadsheet file
        input: PathBuf,
    },

    /// List all sheets in a workbook
    Sheets {
        /// Input spreadsheet file
        input: PathBuf,
    },
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    match cli.command {
        Commands::ToCsv {
            input,
            output,
            sheet,
            calculate,
            delimiter,
        } => to_csv(&input, output.as_deref(), sheet, calculate, delimiter),
        Commands::Info { input } => show_info(&input),
        Commands::Sheets { input } => list_sheets(&input),
    }
}

fn to_csv(
    input: &PathBuf,
    output: Option<&std::path::Path>,
    sheet_idx: usize,
    calculate: bool,
    delimiter: char,
) -> Result<()> {
    // Load the workbook
    let mut workbook =
        Workbook::open(input).with_context(|| format!("Failed to open '{}'", input.display()))?;

    // Calculate formulas if requested
    if calculate {
        let stats = workbook
            .calculate_with_options(&CalculationOptions {
                force_full_calculation: true,
                ..Default::default()
            })
            .context("Failed to calculate formulas")?;

        eprintln!(
            "Calculated {} formulas ({} errors)",
            stats.cells_calculated, stats.errors
        );
    }

    // Get the worksheet
    let sheet = workbook
        .worksheet(sheet_idx)
        .with_context(|| format!("Sheet index {} not found", sheet_idx))?;

    // Get the used range
    let used_range = match sheet.used_range() {
        Some(range) => range,
        None => {
            eprintln!("Warning: Sheet appears to be empty");
            return Ok(());
        }
    };

    let max_row = used_range.end.row;
    let max_col = used_range.end.col;

    // Build CSV output
    let mut csv_output = String::new();

    for row in 0..=max_row {
        let mut first = true;
        for col in 0..=max_col {
            if !first {
                csv_output.push(delimiter);
            }
            first = false;

            // Get cell value (use calculated value if available)
            let value = if calculate {
                sheet.get_calculated_value_at(row, col)
            } else {
                Some(&sheet.get_value_at(row, col))
            };

            if let Some(val) = value {
                let text = cell_value_to_csv_string(val, delimiter);
                csv_output.push_str(&text);
            }
        }
        csv_output.push('\n');
    }

    // Output
    if let Some(output_path) = output {
        std::fs::write(output_path, &csv_output)
            .with_context(|| format!("Failed to write '{}'", output_path.display()))?;
        eprintln!("Wrote {} rows to '{}'", max_row + 1, output_path.display());
    } else {
        io::stdout()
            .write_all(csv_output.as_bytes())
            .context("Failed to write to stdout")?;
    }

    Ok(())
}

/// Convert a CellValue to a CSV-safe string
fn cell_value_to_csv_string(value: &CellValue, delimiter: char) -> String {
    let text = match value {
        CellValue::Empty => String::new(),
        CellValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        CellValue::String(s) => s.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::Error(e) => e.to_string(),
        CellValue::Formula { cached_value, .. } => {
            if let Some(v) = cached_value {
                return cell_value_to_csv_string(v, delimiter);
            }
            String::new()
        }
        CellValue::SpillTarget { .. } => {
            // SpillTarget cells would need to look up the source formula's array result
            // For CSV export, we output empty for now
            String::new()
        }
    };

    // Quote if necessary
    if text.contains(delimiter) || text.contains('"') || text.contains('\n') || text.contains('\r')
    {
        format!("\"{}\"", text.replace('"', "\"\""))
    } else {
        text
    }
}

fn show_info(input: &PathBuf) -> Result<()> {
    let workbook =
        Workbook::open(input).with_context(|| format!("Failed to open '{}'", input.display()))?;

    println!("File: {}", input.display());
    println!("Sheets: {}", workbook.sheet_count());

    for i in 0..workbook.sheet_count() {
        if let Some(sheet) = workbook.worksheet(i) {
            let formula_count = sheet.formula_cells().count();

            println!();
            println!("  Sheet {}: \"{}\"", i, sheet.name());

            if let Some(range) = sheet.used_range() {
                println!(
                    "    Used range: {} rows x {} columns",
                    range.end.row + 1,
                    range.end.col + 1
                );
            } else {
                println!("    Used range: empty");
            }
            println!("    Formulas: {}", formula_count);
        }
    }

    Ok(())
}

fn list_sheets(input: &PathBuf) -> Result<()> {
    let workbook =
        Workbook::open(input).with_context(|| format!("Failed to open '{}'", input.display()))?;

    for i in 0..workbook.sheet_count() {
        if let Some(sheet) = workbook.worksheet(i) {
            println!("{}\t{}", i, sheet.name());
        }
    }

    Ok(())
}
