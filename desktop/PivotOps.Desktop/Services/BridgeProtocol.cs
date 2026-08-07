using System.Text.Json;
using System.Text.Json.Serialization;

namespace PivotOps.Desktop.Services;

/// <summary>Envelope sent by the web app over <c>chrome.webview.postMessage</c>.</summary>
internal sealed class HostRequest
{
    public int Id { get; set; }
    public string Kind { get; set; } = "";
    public JsonElement Payload { get; set; }
}

internal sealed class HostResponse
{
    public int Id { get; set; }
    public bool Ok { get; set; }
    public object? Result { get; set; }
    public string? Error { get; set; }
}

internal sealed class ExcelSyncPayload
{
    public List<ExcelOp> Ops { get; set; } = new();
}

internal sealed class StoragePayload
{
    public string Key { get; set; } = "";
    public string? Value { get; set; }
}

/// <summary>A single queued workbook operation emitted by the Office.js shim.</summary>
internal sealed class ExcelOp
{
    public string T { get; set; } = "";
    public string? Sheet { get; set; }
    public string? Key { get; set; }
    public string? Id { get; set; }
    public string? Target { get; set; }
    public string? Prop { get; set; }
    public string? Source { get; set; }
    public int? R { get; set; }
    public int? C { get; set; }
    public int? Rows { get; set; }
    public int? Cols { get; set; }
    public int? Count { get; set; }
    public bool? HasHeaders { get; set; }
    public bool? InCellDropDown { get; set; }
    public JsonElement Value { get; set; }
    public JsonElement Values { get; set; }
    public JsonElement Formats { get; set; }
    public Dictionary<string, JsonElement>? Props { get; set; }
}

internal static class BridgeJson
{
    public static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };
}
