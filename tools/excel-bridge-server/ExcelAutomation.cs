// Generic COM object store and navigator.
//
// This is NOT Excel-specific. It manages a table of COM object handles and
// provides Get/Set/Invoke operations that navigate chains of properties.
// The entire file is ~100 lines because C#'s `dynamic` handles all the
// IDispatch late-binding automatically.

using System.Text.Json;

namespace ExcelBridgeServer;

/// <summary>
/// Manages a handle table of COM objects and provides generic navigation.
/// </summary>
public sealed class ComObjectStore : IDisposable
{
    private readonly Dictionary<ulong, dynamic> _handles = new();
    private ulong _nextHandle = 1; // 0 is reserved for Excel.Application
    private bool _disposed;

    /// <summary>
    /// Initialize Excel.Application and store as handle 0.
    /// </summary>
    public void InitExcel()
    {
        if (_handles.ContainsKey(0)) return;

        var type = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel.Application not found. Is Excel installed?");
        dynamic app = Activator.CreateInstance(type)!;
        app.Visible = false;
        app.DisplayAlerts = false;
        app.ScreenUpdating = false;
        _handles[0] = app;
        Console.Error.WriteLine("[excel-bridge] Excel.Application created (handle 0)");
    }

    /// <summary>
    /// Navigate a chain of property accesses from a stored handle.
    /// Returns the final COM object at the end of the chain.
    /// </summary>
    public dynamic Navigate(ulong handle, JsonElement? chainEl)
    {
        if (!_handles.TryGetValue(handle, out var obj))
            throw new KeyNotFoundException($"Unknown handle: {handle}");

        if (chainEl == null || chainEl.Value.ValueKind != JsonValueKind.Array)
            return obj;

        foreach (var step in chainEl.Value.EnumerateArray())
        {
            if (step.ValueKind == JsonValueKind.String)
            {
                // Simple property: "Workbooks"
                string prop = step.GetString()!;
                obj = GetDynamicProperty(obj, prop);
            }
            else if (step.ValueKind == JsonValueKind.Array)
            {
                // Indexed property: ["Worksheets", 1] or ["Range", "A1"]
                var arr = step.EnumerateArray().ToList();
                if (arr.Count < 2)
                    throw new ArgumentException($"Indexed chain step needs [name, index], got {step}");
                string prop = arr[0].GetString()!;
                object? index = ProtocolHelpers.JsonToComValue(arr[1]);
                obj = GetDynamicIndexed(obj, prop, index);
            }
            else
            {
                throw new ArgumentException($"Invalid chain step: {step}");
            }
        }

        return obj;
    }

    /// <summary>
    /// Get a property. If the result is a COM object, store it and return a handle.
    /// </summary>
    public (bool isHandle, ulong handle, object? value) GetProperty(dynamic target, string property)
    {
        dynamic result = GetDynamicProperty(target, property);
        return WrapResult(result);
    }

    /// <summary>
    /// Set a property.
    /// </summary>
    public void SetProperty(dynamic target, string property, object? value)
    {
        // Use reflection to set the property dynamically
        // The `dynamic` keyword handles IDispatch property put automatically
        var type = (Type)target.GetType();
        type.InvokeMember(property,
            System.Reflection.BindingFlags.SetProperty,
            null, target, new[] { value });
    }

    /// <summary>
    /// Invoke a method. If the result is a COM object, store it and return a handle.
    /// </summary>
    public (bool isHandle, ulong handle, object? value) InvokeMethod(
        dynamic target, string method, object?[] args)
    {
        var type = (Type)target.GetType();
        var result = type.InvokeMember(method,
            System.Reflection.BindingFlags.InvokeMethod,
            null, target, args);
        return WrapResult(result);
    }

    /// <summary>
    /// Release a handle, freeing the COM reference.
    /// Handle 0 (Excel.Application) cannot be released — use Dispose().
    /// </summary>
    public void Release(ulong handle)
    {
        if (handle == 0)
            throw new InvalidOperationException("Cannot release handle 0. Use Shutdown.");
        if (_handles.Remove(handle, out var obj))
        {
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); }
            catch { /* best effort */ }
        }
    }

    // -- Internals --

    private (bool isHandle, ulong handle, object? value) WrapResult(object? result)
    {
        if (result == null || result is DBNull)
            return (false, 0, null);

        // Check if result is a COM object (has a __ComObject type)
        if (result.GetType().IsCOMObject)
        {
            var h = _nextHandle++;
            _handles[h] = result;
            return (true, h, null);
        }

        return (false, 0, ProtocolHelpers.ComValueToJson(result));
    }

    private static dynamic GetDynamicProperty(dynamic obj, string name)
    {
        var type = (Type)obj.GetType();
        return type.InvokeMember(name,
            System.Reflection.BindingFlags.GetProperty,
            null, obj, null);
    }

    private static dynamic GetDynamicIndexed(dynamic obj, string name, object? index)
    {
        var type = (Type)obj.GetType();
        return type.InvokeMember(name,
            System.Reflection.BindingFlags.GetProperty,
            null, obj, new[] { index });
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        // Quit Excel if it was initialized
        if (_handles.TryGetValue(0, out var app))
        {
            try { app.Quit(); Console.Error.WriteLine("[excel-bridge] Excel quit"); }
            catch (Exception ex) { Console.Error.WriteLine($"[excel-bridge] Quit error: {ex.Message}"); }
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(app); }
            catch { }
        }

        // Release all other handles
        foreach (var kv in _handles.Where(kv => kv.Key != 0))
        {
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(kv.Value); }
            catch { }
        }
        _handles.Clear();
    }
}
