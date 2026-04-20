// Excel COM Bridge Server - Generic COM Proxy
//
// A TCP server that navigates COM objects via Get/Set/Invoke commands.
// All Excel-specific knowledge lives in the client; this server is a
// thin, stable proxy that never needs modification for new features.
//
// Usage:
//   ExcelBridgeServer.exe [--port 9876]

using System.Net;
using System.Net.Sockets;
using System.Text.Json;
using ExcelBridgeServer;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        var port = 9876;
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i] == "--port" && int.TryParse(args[i + 1], out var p))
                port = p;
        }

        Console.Error.WriteLine($"[excel-bridge] Starting on port {port}...");
        var listener = new TcpListener(IPAddress.Any, port);
        listener.Start();
        Console.Error.WriteLine($"[excel-bridge] Listening on 0.0.0.0:{port}");

        // Accept connections in a loop. Each disconnection cleans up Excel,
        // so the VM can stay running across multiple test runs.
        while (true)
        {
            Console.Error.WriteLine("[excel-bridge] Waiting for connection...");
            using var client = listener.AcceptTcpClient();
            client.NoDelay = true;
            Console.Error.WriteLine($"[excel-bridge] Connected: {client.Client.RemoteEndPoint}");

            using var store = new ComObjectStore();

            try
            {
                using var stream = client.GetStream();
                using var reader = new StreamReader(stream);
                using var writer = new StreamWriter(stream) { AutoFlush = true };

                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (string.IsNullOrEmpty(line)) continue;

                    Response response;
                    ulong reqId = 0;
                    bool shutdown = false;

                    try
                    {
                        var req = ProtocolHelpers.ParseRequest(line);
                        reqId = req.Id;
                        (response, shutdown) = Dispatch(store, req);
                    }
                    catch (JsonException ex)
                    {
                        Console.Error.WriteLine($"[excel-bridge] Parse error: {ex.Message}");
                        response = Response.Error(0, $"JSON parse error: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        response = Response.Error(reqId, ex.Message);
                    }

                    writer.WriteLine(ProtocolHelpers.Serialize(response));

                    if (shutdown) break;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[excel-bridge] Connection error: {ex.Message}");
            }

            Console.Error.WriteLine("[excel-bridge] Client disconnected.");
        }
    }

    // -----------------------------------------------------------------------
    // Command dispatch - just 6 cases, all generic
    // -----------------------------------------------------------------------

    static (Response resp, bool shutdown) Dispatch(ComObjectStore store, Request req)
    {
        var id = req.Id;
        var p = req.Params;

        try
        {
            switch (req.Cmd)
            {
                case "Init":
                    store.InitExcel();
                    return (Response.Ok(id), false);

                case "Get":
                {
                    var handle = p!.Value.GetProperty("handle").GetUInt64();
                    var chain = p.Value.TryGetProperty("chain", out var c) ? (JsonElement?)c : null;
                    var property = p.Value.GetProperty("property").GetString()!;

                    var target = store.Navigate(handle, chain);
                    var result = ((bool isHandle, ulong handle, object? value))
                        store.GetProperty(target, property);

                    object? data = result.isHandle ? new HandleData(result.handle) : new ValueData(result.value);
                    return (Response.Ok(id, data), false);
                }

                case "Set":
                {
                    var handle = p!.Value.GetProperty("handle").GetUInt64();
                    var chain = p.Value.TryGetProperty("chain", out var c) ? (JsonElement?)c : null;
                    var property = p.Value.GetProperty("property").GetString()!;
                    var value = ProtocolHelpers.JsonToComValue(p.Value.GetProperty("value"));

                    var target = store.Navigate(handle, chain);
                    store.SetProperty(target, property, value);
                    return (Response.Ok(id), false);
                }

                case "Invoke":
                {
                    var handle = p!.Value.GetProperty("handle").GetUInt64();
                    var chain = p.Value.TryGetProperty("chain", out var c) ? (JsonElement?)c : null;
                    var method = p.Value.GetProperty("method").GetString()!;

                    object?[] invokeArgs = Array.Empty<object?>();
                    if (p.Value.TryGetProperty("args", out var argsEl)
                        && argsEl.ValueKind == JsonValueKind.Array)
                    {
                        invokeArgs = argsEl.EnumerateArray()
                            .Select(el => {
                                // Handle references: {"$ref": <handle_id>} resolves to stored COM object
                                if (el.ValueKind == JsonValueKind.Object
                                    && el.TryGetProperty("$ref", out var refEl)
                                    && refEl.ValueKind == JsonValueKind.Number)
                                {
                                    return (object?)store.Navigate(refEl.GetUInt64(), null);
                                }
                                return ProtocolHelpers.JsonToComValue(el);
                            })
                            .Select(a => a ?? System.Reflection.Missing.Value)
                            .ToArray();
                    }

                    var target = store.Navigate(handle, chain);
                    var result = ((bool isHandle, ulong handle, object? value))
                        store.InvokeMethod(target, method, invokeArgs);

                    object? data = result.isHandle ? new HandleData(result.handle) : result.value != null ? new ValueData(result.value) : null;
                    return (Response.Ok(id, data), false);
                }

                case "Navigate":
                {
                    var handle = p!.Value.GetProperty("handle").GetUInt64();
                    var chain = p.Value.TryGetProperty("chain", out var c) ? (JsonElement?)c : null;
                    var target = store.Navigate(handle, chain);
                    var result = store.StoreAndReturnHandle(target);
                    return (Response.Ok(id, new HandleData(result)), false);
                }


                case "Release":
                {
                    var handle = p!.Value.GetProperty("handle").GetUInt64();
                    store.Release(handle);
                    return (Response.Ok(id), false);
                }

                case "Shutdown":
                    store.Dispose();
                    return (Response.Ok(id), true);

                default:
                    return (Response.Error(id, $"Unknown command: {req.Cmd}"), false);
            }
        }
        catch (Exception ex)
        {
            var msg = ex.InnerException != null
                ? $"{req.Cmd}: {ex.InnerException.Message}"
                : $"{req.Cmd}: {ex.Message}";
            return (Response.Error(id, msg), false);
        }
    }
}
