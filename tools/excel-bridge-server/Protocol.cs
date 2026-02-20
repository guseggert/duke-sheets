// Protocol types for the generic COM proxy.
//
// Wire format: newline-delimited JSON (NDJSON), one object per line.
//
// The protocol has only 5 commands: Init, Get, Set, Invoke, Release, Shutdown.
// All Excel-specific knowledge lives in the client — the server is a thin
// COM object navigator.

using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelBridgeServer;

// ---------------------------------------------------------------------------
// Request (parsed manually due to flattened cmd/params layout)
// ---------------------------------------------------------------------------

public record Request(ulong Id, string Cmd, JsonElement? Params);

// ---------------------------------------------------------------------------
// Response
// ---------------------------------------------------------------------------

public class Response
{
    [JsonPropertyName("id")]
    public ulong Id { get; set; }

    [JsonPropertyName("status")]
    public string Status { get; set; } = "ok";

    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonPropertyName("data")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Data { get; set; }

    public static Response Ok(ulong id, object? data = null) =>
        new() { Id = id, Status = "ok", Data = data };

    public static Response Error(ulong id, string message) =>
        new() { Id = id, Status = "error", Message = message };
}

// Response data shapes
public record HandleData([property: JsonPropertyName("handle")] ulong Handle);
public record ValueData([property: JsonPropertyName("value")] object? Value);

// ---------------------------------------------------------------------------
// Parsing helpers
// ---------------------------------------------------------------------------

public static class ProtocolHelpers
{
    public static Request ParseRequest(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        var id = root.GetProperty("id").GetUInt64();
        var cmd = root.GetProperty("cmd").GetString()
                  ?? throw new JsonException("Missing 'cmd'");
        JsonElement? parms = root.TryGetProperty("params", out var p) ? p.Clone() : null;
        return new Request(id, cmd, parms);
    }

    /// <summary>
    /// Convert a JSON value to a .NET object suitable for COM dispatch.
    /// </summary>
    public static object? JsonToComValue(JsonElement el)
    {
        return el.ValueKind switch
        {
            JsonValueKind.Null => null,
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number => el.GetDouble(),
            JsonValueKind.String => el.GetString(),
            _ => null,
        };
    }

    /// <summary>
    /// Convert a COM return value to a JSON-compatible .NET object.
    /// </summary>
    public static object? ComValueToJson(object? value)
    {
        if (value == null || value is DBNull) return null;
        if (value is bool b) return b;
        if (value is double d) return d;
        if (value is int i) return (double)i;
        if (value is float f) return (double)f;
        if (value is decimal dec) return (double)dec;
        if (value is string s) return s;
        return value.ToString();
    }

    /// <summary>
    /// Serialize a Response to a single JSON line.
    /// </summary>
    public static string Serialize(Response resp)
    {
        return JsonSerializer.Serialize(resp, JsonCtx.Default.Response);
    }
}

[JsonSerializable(typeof(Response))]
[JsonSerializable(typeof(HandleData))]
[JsonSerializable(typeof(ValueData))]
internal partial class JsonCtx : JsonSerializerContext { }
