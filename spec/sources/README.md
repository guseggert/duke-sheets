# Spec Sources

This directory holds committed copies of the format specifications that
drive the feature taxonomy. Files here are **read-only inputs** -
`tools/generate_taxonomy.py` parses them to regenerate
`spec/features/taxonomy.toml` and `crates/duke-sheets-features/src/generated.rs`.

Committing the sources rather than fetching on demand means:

- Regeneration is reproducible offline.
- Spec version is pinned with the repo.
- CI doesn't depend on external sites being reachable.

## Provenance

| File | Source | Format | Used by |
|------|--------|--------|---------|
| `ecma-376/sml.xsd` | [ECMA-376 5th edition Part 1, Annex A](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) | XSD | `generate_taxonomy/xsd_xlsx.py` |
| `ms-xlsb/records-by-number.html` | [[MS-XLSB] §2.1.2 Records By Number](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/) | HTML | `generate_taxonomy/msxlsb.py` |
| `ms-xlsb/sections/*.html` | [[MS-XLSB] §2.4.x Record Definitions](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/) | HTML | `generate_taxonomy/msxlsb.py` |
| `ms-xls/records-by-number.html` | [[MS-XLS] §2.1.2 Records By Number](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/) | HTML | `generate_taxonomy/msxls.py` |
| `ms-xls/sections/*.html` | [[MS-XLS] §2.4.x Record Definitions](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/) | HTML | `generate_taxonomy/msxls.py` |
| `odf-1.3/OpenDocument-schema.rng` | [ODF 1.3 Part 3: Packages](https://docs.oasis-open.org/office/OpenDocument/v1.3/os/part3-schema/) | RelaxNG | `generate_taxonomy/rng_ods.py` |

## Acquiring sources

Sources are acquired manually when the spec version changes (rare - these
specs are effectively frozen). There is no automated fetch step.

When adding or updating a source:

1. Download from the canonical location above.
2. Commit alongside a brief note in this README if the filename or
   structure changes.
3. Run `mise run features:generate` to regenerate downstream artifacts.
4. Run `mise run features:check` to validate no tags drifted.
