//! Additional read-only methods for the Workbook class.

use napi::bindgen_prelude::*;
use napi_derive::napi;

use duke_sheets_core::named_range::NameScope;

use super::{catch_panic, to_napi_err, JsChartSheet, JsNamedRange, JsSheetSlot, JsWorkbookSettings, Workbook};

#[napi]
impl Workbook {
    /// Whether the workbook has no worksheets.
    #[napi(getter)]
    pub fn is_empty(&self) -> Result<bool> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.is_empty())
        })
    }

    /// The index of the active (selected) worksheet.
    #[napi(getter)]
    pub fn active_sheet(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.active_sheet() as u32)
        })
    }

    /// Get the index of a worksheet by name, or null if not found.
    #[napi]
    pub fn sheet_index(&self, name: String) -> Result<Option<u32>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.sheet_index(&name).map(|i| i as u32))
        })
    }

    /// Workbook-level settings (date system, protection, etc.).
    #[napi(getter)]
    pub fn settings(&self) -> Result<JsWorkbookSettings> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(JsWorkbookSettings::from(wb.settings()))
        })
    }

    /// Get all named ranges defined in the workbook.
    #[napi(getter)]
    pub fn named_ranges(&self) -> Result<Vec<JsNamedRange>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb
                .named_ranges()
                .iter()
                .map(|nr| JsNamedRange {
                    name: nr.name.clone(),
                    scope: match &nr.scope {
                        NameScope::Workbook => "workbook".into(),
                        NameScope::Sheet(_) => "sheet".into(),
                    },
                    sheet_index: match &nr.scope {
                        NameScope::Workbook => None,
                        NameScope::Sheet(idx) => Some(*idx as u32),
                    },
                    refers_to: nr.refers_to.clone(),
                    comment: nr.comment.clone(),
                    hidden: nr.hidden,
                })
                .collect())
        })
    }
}

#[napi]
impl Workbook {
    /// Get all chart sheets.
    #[napi(getter)]
    pub fn chartsheets(&self) -> Result<Vec<JsChartSheet>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.chartsheets().iter().map(JsChartSheet::from).collect())
        })
    }

    /// Get the number of chart sheets.
    #[napi(getter)]
    pub fn chartsheet_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.chartsheet_count() as u32)
        })
    }

    /// Get the tab-bar ordering of worksheets and chartsheets.
    #[napi(getter)]
    pub fn sheet_order(&self) -> Result<Vec<JsSheetSlot>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.sheet_order().iter().map(JsSheetSlot::from).collect())
        })
    }

    /// Total number of tabs (worksheets + chartsheets).
    #[napi(getter)]
    pub fn total_sheet_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.total_sheet_count() as u32)
        })
    }
}
