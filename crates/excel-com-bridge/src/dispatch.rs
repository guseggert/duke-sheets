//! Safe wrapper around IDispatch for late-bound COM automation.
//!
//! Excel's COM API is primarily accessed through IDispatch (like VBScript late-binding).
//! This module provides ergonomic helpers for property get/set and method invocation.

#![cfg(windows)]

use std::mem::ManuallyDrop;
use std::ptr;

use windows::{
    core::{BSTR, GUID, HSTRING, PCWSTR},
    Win32::{
        Foundation::{DISP_E_EXCEPTION, VARIANT_BOOL},
        Globalization::GetSystemDefaultLCID,
        System::{
            Com::{
                CLSIDFromProgID, CoCreateInstance, IDispatch, CLSCTX_LOCAL_SERVER, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
            },
            Ole::DISPID_PROPERTYPUT,
            Variant::{
                VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_ERROR, VT_I2, VT_I4, VT_NULL,
                VT_R4, VT_R8,
            },
        },
    },
};

// -- VARIANT construction helpers --
// The VARIANT struct wraps inner unions in ManuallyDrop, so we use ptr::write
// to set fields without triggering the DerefMut lint.

/// Create an empty VARIANT.
pub fn variant_empty() -> VARIANT {
    VARIANT::default()
}

/// Create a VARIANT containing a bool.
pub fn variant_bool(val: bool) -> VARIANT {
    unsafe {
        let mut v = VARIANT::default();
        let inner = &mut *v.Anonymous.Anonymous;
        ptr::write(&mut inner.vt, VT_BOOL);
        ptr::write(
            &mut inner.Anonymous.boolVal,
            VARIANT_BOOL(if val { -1 } else { 0 }),
        );
        v
    }
}

/// Create a VARIANT containing an f64.
pub fn variant_f64(val: f64) -> VARIANT {
    unsafe {
        let mut v = VARIANT::default();
        let inner = &mut *v.Anonymous.Anonymous;
        ptr::write(&mut inner.vt, VT_R8);
        ptr::write(&mut inner.Anonymous.dblVal, val);
        v
    }
}

/// Create a VARIANT containing an i32.
pub fn variant_i32(val: i32) -> VARIANT {
    unsafe {
        let mut v = VARIANT::default();
        let inner = &mut *v.Anonymous.Anonymous;
        ptr::write(&mut inner.vt, VT_I4);
        ptr::write(&mut inner.Anonymous.lVal, val);
        v
    }
}

/// Create a VARIANT containing a BSTR string.
pub fn variant_str(val: &str) -> VARIANT {
    unsafe {
        let bstr = BSTR::from(val);
        let mut v = VARIANT::default();
        let inner = &mut *v.Anonymous.Anonymous;
        ptr::write(&mut inner.vt, VT_BSTR);
        ptr::write(&mut inner.Anonymous.bstrVal, ManuallyDrop::new(bstr));
        v
    }
}

/// Get the VT type of a VARIANT.
pub fn variant_vt(v: &VARIANT) -> u16 {
    unsafe { v.Anonymous.Anonymous.vt.0 }
}

/// Extract a bool from a VARIANT.
pub fn variant_get_bool(v: &VARIANT) -> Option<bool> {
    unsafe {
        if v.Anonymous.Anonymous.vt == VT_BOOL {
            Some(v.Anonymous.Anonymous.Anonymous.boolVal.0 != 0)
        } else {
            None
        }
    }
}

/// Extract an f64 from a VARIANT.
pub fn variant_get_f64(v: &VARIANT) -> Option<f64> {
    unsafe {
        let vt = v.Anonymous.Anonymous.vt;
        let anon = &v.Anonymous.Anonymous.Anonymous;
        if vt == VT_R8 {
            Some(anon.dblVal)
        } else if vt == VT_R4 {
            Some(anon.fltVal as f64)
        } else if vt == VT_I4 {
            Some(anon.lVal as f64)
        } else if vt == VT_I2 {
            Some(anon.iVal as f64)
        } else {
            None
        }
    }
}

/// Extract a string from a VARIANT.
pub fn variant_get_string(v: &VARIANT) -> Option<String> {
    unsafe {
        if v.Anonymous.Anonymous.vt == VT_BSTR {
            let bstr = &v.Anonymous.Anonymous.Anonymous.bstrVal;
            Some(bstr.to_string())
        } else {
            None
        }
    }
}

/// Extract an IDispatch from a VARIANT.
pub fn variant_get_dispatch(v: &VARIANT) -> Option<IDispatch> {
    unsafe {
        if v.Anonymous.Anonymous.vt == VT_DISPATCH {
            // pdispVal is ManuallyDrop<Option<IDispatch>>
            let opt_disp: &Option<IDispatch> = &v.Anonymous.Anonymous.Anonymous.pdispVal;
            opt_disp.clone()
        } else {
            None
        }
    }
}

/// Check if a VARIANT is empty or null.
pub fn variant_is_empty(v: &VARIANT) -> bool {
    unsafe {
        let vt = v.Anonymous.Anonymous.vt;
        vt == VT_EMPTY || vt == VT_NULL
    }
}

/// Check if a VARIANT is a VT_ERROR.
pub fn variant_is_error(v: &VARIANT) -> bool {
    unsafe { v.Anonymous.Anonymous.vt == VT_ERROR }
}

// -- DispatchObject --

/// A wrapper around an IDispatch COM object providing ergonomic access.
#[derive(Clone)]
pub struct DispatchObject {
    inner: IDispatch,
}

impl DispatchObject {
    /// Create a COM object from a ProgID string (e.g., "Excel.Application").
    pub fn create_from_progid(progid: &str) -> Result<Self, String> {
        unsafe {
            let hstr = HSTRING::from(progid);
            let clsid =
                CLSIDFromProgID(&hstr).map_err(|e| format!("CLSIDFromProgID failed: {e}"))?;
            let disp: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)
                .map_err(|e| format!("CoCreateInstance failed for '{progid}': {e}"))?;
            Ok(Self { inner: disp })
        }
    }

    /// Wrap an existing IDispatch pointer.
    pub fn from_idispatch(disp: IDispatch) -> Self {
        Self { inner: disp }
    }

    /// Look up the DISPID for a member name.
    fn get_dispid(&self, name: &str) -> Result<i32, String> {
        unsafe {
            let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
            let pcwstr = PCWSTR(wide.as_ptr());
            let names = [pcwstr];
            let mut dispid = 0i32;
            self.inner
                .GetIDsOfNames(
                    &GUID::zeroed(),
                    names.as_ptr(),
                    1,
                    GetSystemDefaultLCID(),
                    &mut dispid,
                )
                .map_err(|e| format!("GetIDsOfNames('{name}') failed: {e}"))?;
            Ok(dispid)
        }
    }

    /// Get a property value. Equivalent to VB's `obj.PropertyName`.
    pub fn get_property(&self, name: &str) -> Result<VARIANT, String> {
        let dispid = self.get_dispid(name)?;
        unsafe {
            let params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            let mut except = EXCEPINFO::default();
            self.inner
                .Invoke(
                    dispid,
                    &GUID::zeroed(),
                    GetSystemDefaultLCID(),
                    DISPATCH_PROPERTYGET,
                    &params,
                    Some(&mut result),
                    Some(&mut except),
                    None,
                )
                .map_err(|e| format_invoke_error(e, &except, name))?;
            Ok(result)
        }
    }

    /// Set a property value. Equivalent to VB's `obj.PropertyName = value`.
    pub fn set_property(&self, name: &str, value: VARIANT) -> Result<(), String> {
        let dispid = self.get_dispid(name)?;
        unsafe {
            let mut args = [value];
            let mut named_args = [DISPID_PROPERTYPUT];
            let params = DISPPARAMS {
                rgvarg: args.as_mut_ptr(),
                rgdispidNamedArgs: named_args.as_mut_ptr(),
                cArgs: 1,
                cNamedArgs: 1,
            };
            let mut except = EXCEPINFO::default();
            self.inner
                .Invoke(
                    dispid,
                    &GUID::zeroed(),
                    GetSystemDefaultLCID(),
                    DISPATCH_PROPERTYPUT,
                    &params,
                    None,
                    Some(&mut except),
                    None,
                )
                .map_err(|e| format_invoke_error(e, &except, name))?;
            Ok(())
        }
    }

    /// Invoke a method with arguments. Arguments should be in natural order
    /// (this function reverses them as required by DISPPARAMS).
    pub fn invoke_method(&self, name: &str, args: &[VARIANT]) -> Result<VARIANT, String> {
        let dispid = self.get_dispid(name)?;
        unsafe {
            // DISPPARAMS requires arguments in reverse order
            let mut reversed: Vec<VARIANT> = args.iter().rev().cloned().collect();
            let params = DISPPARAMS {
                rgvarg: if reversed.is_empty() {
                    std::ptr::null_mut()
                } else {
                    reversed.as_mut_ptr()
                },
                rgdispidNamedArgs: std::ptr::null_mut(),
                cArgs: reversed.len() as u32,
                cNamedArgs: 0,
            };
            let mut result = VARIANT::default();
            let mut except = EXCEPINFO::default();
            self.inner
                .Invoke(
                    dispid,
                    &GUID::zeroed(),
                    GetSystemDefaultLCID(),
                    DISPATCH_METHOD,
                    &params,
                    Some(&mut result),
                    Some(&mut except),
                    None,
                )
                .map_err(|e| format_invoke_error(e, &except, name))?;
            Ok(result)
        }
    }

    /// Get a child object (property that returns an IDispatch).
    pub fn get_child(&self, name: &str) -> Result<DispatchObject, String> {
        let variant = self.get_property(name)?;
        extract_dispatch(&variant, name)
    }

    /// Invoke a method and extract the returned IDispatch object.
    pub fn invoke_child(&self, name: &str, args: &[VARIANT]) -> Result<DispatchObject, String> {
        let variant = self.invoke_method(name, args)?;
        extract_dispatch(&variant, name)
    }

    /// Get a property that's indexed (e.g., `Worksheets(1)` or `Range("A1")`).
    pub fn get_indexed(&self, name: &str, index: &VARIANT) -> Result<DispatchObject, String> {
        let dispid = self.get_dispid(name)?;
        unsafe {
            let mut args = [index.clone()];
            let params = DISPPARAMS {
                rgvarg: args.as_mut_ptr(),
                rgdispidNamedArgs: std::ptr::null_mut(),
                cArgs: 1,
                cNamedArgs: 0,
            };
            let mut result = VARIANT::default();
            let mut except = EXCEPINFO::default();
            self.inner
                .Invoke(
                    dispid,
                    &GUID::zeroed(),
                    GetSystemDefaultLCID(),
                    DISPATCH_PROPERTYGET,
                    &params,
                    Some(&mut result),
                    Some(&mut except),
                    None,
                )
                .map_err(|e| format_invoke_error(e, &except, name))?;
            extract_dispatch(&result, name)
        }
    }
}

/// Extract an IDispatch from a VARIANT, with a descriptive error.
fn extract_dispatch(variant: &VARIANT, context: &str) -> Result<DispatchObject, String> {
    if let Some(disp) = variant_get_dispatch(variant) {
        Ok(DispatchObject::from_idispatch(disp))
    } else if variant_is_empty(variant) {
        Err(format!("'{context}' returned empty/null"))
    } else {
        let vt = variant_vt(variant);
        Err(format!(
            "'{context}' returned non-object VARIANT (VT={vt}), expected VT_DISPATCH"
        ))
    }
}

/// Format an Invoke error, including EXCEPINFO details if available.
fn format_invoke_error(err: windows::core::Error, except: &EXCEPINFO, member_name: &str) -> String {
    let code = err.code().0 as u32;
    if code == DISP_E_EXCEPTION.0 as u32 {
        let desc = if !except.bstrDescription.is_empty() {
            except.bstrDescription.to_string()
        } else {
            String::from("(no description)")
        };
        let source = if !except.bstrSource.is_empty() {
            except.bstrSource.to_string()
        } else {
            String::from("(no source)")
        };
        format!("COM exception in '{member_name}': {desc} (source: {source})")
    } else {
        format!("Invoke('{member_name}') failed: {err}")
    }
}
