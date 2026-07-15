//! Additional read-only methods for the Workbook class.

use napi::bindgen_prelude::*;
use napi_derive::napi;

use duke_sheets_core::named_range::NameScope;

use super::{
    catch_panic, to_napi_err, JsChartSheet, JsNamedRange, JsSheetSlot, JsWorkbookProtection,
    JsWorkbookSettings, Workbook,
};

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

    /// Workbook structure/window protection settings, or null if unprotected.
    #[napi(getter)]
    pub fn workbook_protection(&self) -> Result<Option<JsWorkbookProtection>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb
                .workbook_protection()
                .map(|protection| JsWorkbookProtection::from(&protection)))
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

    /// The workbook theme's 12 clrScheme colors as `RRGGBB` hex, in
    /// theme-index order (background 1, text 1, background 2, text 2,
    /// accent 1-6, hyperlink, followed hyperlink). The Office default
    /// palette when the file carries no theme.
    #[napi(getter)]
    pub fn theme_palette(&self) -> Result<Vec<String>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb
                .theme_palette()
                .colors
                .iter()
                .map(|(r, g, b)| format!("{r:02X}{g:02X}{b:02X}"))
                .collect())
        })
    }

    /// Resolve a drawing color to display RGB (`RRGGBB` hex) against
    /// this workbook's theme palette. `auto` has no fixed RGB and
    /// resolves to `null`.
    #[napi(ts_args_type = "color: object", ts_return_type = "string | null")]
    pub fn resolve_color(&self, env: Env, color: Unknown) -> Result<Option<String>> {
        catch_panic(|| {
            let color = crate::drawings::drawing_color_from_js(&env, color)?;
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb
                .resolve_color(&color)
                .map(|(r, g, b)| format!("{r:02X}{g:02X}{b:02X}")))
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
