//! Fuzz target for the BIFF8 formula token parser + decompiler.
//!
//! Uses `Arbitrary` to generate structured token sequences and
//! decompilation contexts, exercising the full parse → decompile
//! pipeline with controlled inputs.

#![no_main]
use arbitrary::{Arbitrary, Unstructured};
use libfuzzer_sys::fuzz_target;

use duke_sheets_xls::biff::formula::decompiler;
use duke_sheets_xls::biff::formula::token_parser;
use duke_sheets_xls::biff::formula::{ExternSheetEntry, FormulaContext, NameRecord, SupBook};

/// Structured input for the BIFF token fuzzer.
#[derive(Debug)]
struct FuzzBiff {
    /// Raw token bytes (the token parser handles arbitrary bytes)
    token_data: Vec<u8>,
    /// Optional extra data section (for tArray constants)
    extra_data: Vec<u8>,
    /// Context for decompilation
    ctx: FuzzContext,
}

impl<'a> Arbitrary<'a> for FuzzBiff {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        // Cap byte vectors to prevent OOM
        let token_len = u.int_in_range(0..=2048)?;
        let mut token_data = Vec::with_capacity(token_len);
        for _ in 0..token_len {
            token_data.push(u.arbitrary()?);
        }
        let extra_len = u.int_in_range(0..=1024)?;
        let mut extra_data = Vec::with_capacity(extra_len);
        for _ in 0..extra_len {
            extra_data.push(u.arbitrary()?);
        }
        Ok(FuzzBiff {
            token_data,
            extra_data,
            ctx: FuzzContext::arbitrary(u)?,
        })
    }
}

#[derive(Debug)]
struct FuzzContext {
    sheet_names: Vec<String>,
    extern_entries: Vec<FuzzExternEntry>,
    supbook_count: u8,
    names: Vec<FuzzName>,
    base_cell: Option<(u16, u16)>,
}

impl<'a> Arbitrary<'a> for FuzzContext {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let nsheets = u.int_in_range(1..=8)?;
        let mut sheet_names = Vec::with_capacity(nsheets);
        for i in 0..nsheets {
            let name = if u.arbitrary()? {
                format!("Sheet{}", i + 1)
            } else {
                let len = u.int_in_range(1..=20)?;
                let mut s = String::with_capacity(len);
                for _ in 0..len {
                    s.push(u.int_in_range(b'A'..=b'z')? as char);
                }
                s
            };
            sheet_names.push(name);
        }

        let nextern = u.int_in_range(0..=8)?;
        let mut extern_entries = Vec::with_capacity(nextern);
        for _ in 0..nextern {
            extern_entries.push(FuzzExternEntry::arbitrary(u)?);
        }

        let supbook_count = u.int_in_range(1..=4)?;

        let nnames = u.int_in_range(0..=4)?;
        let mut names = Vec::with_capacity(nnames);
        for _ in 0..nnames {
            names.push(FuzzName::arbitrary(u)?);
        }

        let base_cell = if u.arbitrary()? {
            Some((u.arbitrary()?, u.arbitrary()?))
        } else {
            None
        };

        Ok(FuzzContext {
            sheet_names,
            extern_entries,
            supbook_count,
            names,
            base_cell,
        })
    }
}

#[derive(Arbitrary, Debug)]
struct FuzzExternEntry {
    sup_book_idx: u16,
    first_sheet: u16,
    last_sheet: u16,
}

#[derive(Debug)]
struct FuzzName {
    name: String,
    sheet_idx: u16,
    is_builtin: bool,
    formula_body: Vec<u8>,
}

impl<'a> Arbitrary<'a> for FuzzName {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let name = if u.arbitrary()? {
            // Use a builtin name
            let builtins = [
                "Print_Area",
                "Print_Titles",
                "_FilterDatabase",
                "Criteria",
                "Database",
            ];
            let idx = u.int_in_range(0..=builtins.len() - 1)?;
            builtins[idx].into()
        } else {
            let len = u.int_in_range(1..=16)?;
            let mut s = String::with_capacity(len);
            for _ in 0..len {
                s.push(u.int_in_range(b'A'..=b'z')? as char);
            }
            s
        };
        let body_len = u.int_in_range(0..=32)?;
        let mut formula_body = Vec::with_capacity(body_len);
        for _ in 0..body_len {
            formula_body.push(u.arbitrary()?);
        }
        Ok(FuzzName {
            name,
            sheet_idx: u.arbitrary()?,
            is_builtin: u.arbitrary()?,
            formula_body,
        })
    }
}

impl FuzzContext {
    fn to_formula_context(&self) -> FormulaContext {
        let extern_sheet: Vec<ExternSheetEntry> = self
            .extern_entries
            .iter()
            .map(|e| ExternSheetEntry {
                sup_book_idx: e.sup_book_idx,
                first_sheet: e.first_sheet,
                last_sheet: e.last_sheet,
            })
            .collect();

        let supbooks: Vec<SupBook> = (0..self.supbook_count)
            .map(|_| SupBook::SelfRef {
                sheet_count: self.sheet_names.len() as u16,
            })
            .collect();

        let names: Vec<NameRecord> = self
            .names
            .iter()
            .map(|n| NameRecord {
                name: n.name.clone(),
                sheet_idx: n.sheet_idx,
                is_builtin: n.is_builtin,
                formula_body: n.formula_body.clone(),
            })
            .collect();

        let mut ctx = FormulaContext::new(self.sheet_names.clone());
        ctx.extern_sheet = extern_sheet;
        ctx.supbooks = supbooks;
        ctx.names = names;
        if let Some((row, col)) = self.base_cell {
            ctx.set_base_cell(row, col);
        }
        ctx
    }
}

fuzz_target!(|data: &[u8]| {
    if let Ok(input) = FuzzBiff::arbitrary(&mut Unstructured::new(data)) {
        // Cap input sizes to avoid OOM
        if input.token_data.len() > 4096 || input.extra_data.len() > 4096 {
            return;
        }

        // Phase 1: Parse tokens (no extra data)
        let tokens = token_parser::parse_tokens(&input.token_data);

        // Phase 2: Parse tokens with extra data
        let tokens_extra =
            token_parser::parse_tokens_with_extra(&input.token_data, &input.extra_data);

        // Phase 3: Decompile both token sets
        let ctx = input.ctx.to_formula_context();
        let _ = decompiler::decompile(&tokens, &ctx);
        let _ = decompiler::decompile(&tokens_extra, &ctx);
    }

    // Also try raw bytes directly (catches low-level parsing issues)
    let _ = token_parser::parse_tokens(data);
});
