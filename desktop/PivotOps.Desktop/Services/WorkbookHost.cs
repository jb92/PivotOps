using System.Text.Json;
using ClosedXML.Excel;

namespace PivotOps.Desktop.Services;

/// <summary>
/// Applies the operation batches emitted by the Office.js shim to a real .xlsx
/// file, replacing the workbook that Excel used to provide to the add-in.
/// </summary>
internal sealed class WorkbookHost
{
    private readonly SemaphoreSlim _gate = new(1, 1);
    private XLWorkbook? _workbook;
    private DateTime _fileStamp;

    public WorkbookHost(string workbookPath)
    {
        WorkbookPath = workbookPath;
        Directory.CreateDirectory(Path.GetDirectoryName(workbookPath)!);
    }

    public string WorkbookPath { get; }

    public bool Exists => File.Exists(WorkbookPath);

    /// <summary>Raised after a batch has been persisted to disk.</summary>
    public event EventHandler<string>? Saved;

    public async Task<Dictionary<string, object?>> ApplyAsync(IReadOnlyList<ExcelOp> ops)
    {
        await _gate.WaitAsync().ConfigureAwait(false);
        try
        {
            var results = await Task.Run(() => Apply(ops)).ConfigureAwait(false);
            return results;
        }
        finally
        {
            _gate.Release();
        }
    }

    private Dictionary<string, object?> Apply(IReadOnlyList<ExcelOp> ops)
    {
        var workbook = EnsureLoaded();
        var results = new Dictionary<string, object?>();
        var newTables = new Dictionary<string, IXLTable>(StringComparer.Ordinal);
        var mutated = false;

        foreach (var op in ops)
        {
            var sheetName = op.Sheet ?? string.Empty;

            switch (op.T)
            {
                case "sheetGetOrNull":
                    results[op.Key!] = !workbook.Worksheets.TryGetWorksheet(sheetName, out _);
                    break;

                case "sheetAdd":
                    if (!workbook.Worksheets.TryGetWorksheet(sheetName, out _))
                    {
                        workbook.Worksheets.Add(sheetName);
                        mutated = true;
                    }
                    break;

                case "activate":
                    if (TryGetSheet(workbook, op, out var toActivate))
                    {
                        toActivate.SetTabActive();
                        mutated = true;
                    }
                    break;

                case "clear":
                    if (TryGetSheet(workbook, op, out var toClear))
                    {
                        if (TryArea(toClear, op, out var clearRange)) clearRange.Clear(XLClearOptions.All);
                        else toClear.Clear(XLClearOptions.All);
                        mutated = true;
                    }
                    break;

                case "tablesLoad":
                    results[op.Key!] = TryGetSheet(workbook, op, out var tableSheet)
                        ? tableSheet.Tables.Select(t => t.Name).ToArray()
                        : Array.Empty<string>();
                    break;

                case "tableDelete":
                    if (TryGetSheet(workbook, op, out var deleteSheet))
                    {
                        var table = ResolveTable(deleteSheet, newTables, op);
                        if (table is not null)
                        {
                            deleteSheet.Tables.Remove(table.Name);
                            mutated = true;
                        }
                    }
                    break;

                case "tableAdd":
                    if (TryGetSheet(workbook, op, out var addSheet) && TryArea(addSheet, op, out var tableArea))
                    {
                        var created = tableArea.CreateTable();
                        if (op.HasHeaders == false) created.SetShowHeaderRow(false);
                        if (op.Id is not null) newTables[op.Id] = created;
                        mutated = true;
                    }
                    break;

                case "tableSet":
                    if (TryGetSheet(workbook, op, out var setSheet))
                    {
                        ApplyTableProperty(ResolveTable(setSheet, newTables, op), op);
                        mutated = true;
                    }
                    break;

                case "setValues":
                    if (TryGetSheet(workbook, op, out var valueSheet))
                    {
                        WriteValues(valueSheet, op);
                        mutated = true;
                    }
                    break;

                case "format":
                    if (TryGetSheet(workbook, op, out var formatSheet))
                    {
                        ApplyFormat(formatSheet, op);
                        mutated = true;
                    }
                    break;

                case "numberFormat":
                    if (TryGetSheet(workbook, op, out var numberSheet) && TryArea(numberSheet, op, out var numberRange))
                    {
                        var pattern = FirstString(op.Formats);
                        if (pattern is not null) numberRange.Style.NumberFormat.Format = pattern;
                        mutated = true;
                    }
                    break;

                case "validationList":
                    if (TryGetSheet(workbook, op, out var dvSheet) && TryArea(dvSheet, op, out var dvRange))
                    {
                        ApplyListValidation(dvRange, op);
                        mutated = true;
                    }
                    break;

                case "freezeRows":
                    if (TryGetSheet(workbook, op, out var freezeSheet))
                    {
                        freezeSheet.SheetView.FreezeRows(op.Count ?? 0);
                        mutated = true;
                    }
                    break;

                case "autofit":
                    if (TryGetSheet(workbook, op, out var fitSheet))
                    {
                        if (op.C is int firstCol && op.Cols is int colSpan && colSpan > 0)
                            fitSheet.Columns(firstCol + 1, firstCol + colSpan).AdjustToContents();
                        else
                            fitSheet.Columns().AdjustToContents();
                        mutated = true;
                    }
                    break;

                case "getValues":
                    results[op.Key!] = TryGetSheet(workbook, op, out var readSheet) ? ReadValues(readSheet, op) : null;
                    break;
            }
        }

        if (mutated) Save(workbook);
        return results;
    }

    private XLWorkbook EnsureLoaded()
    {
        if (File.Exists(WorkbookPath))
        {
            // Pick up edits the user made in Excel since the last batch.
            var stamp = File.GetLastWriteTimeUtc(WorkbookPath);
            if (_workbook is null || stamp != _fileStamp)
            {
                _workbook?.Dispose();
                _workbook = new XLWorkbook(WorkbookPath);
                _fileStamp = stamp;
            }
        }
        else
        {
            _workbook?.Dispose();
            _workbook = new XLWorkbook();
        }

        return _workbook;
    }

    private void Save(XLWorkbook workbook)
    {
        if (!workbook.Worksheets.Any()) return;
        workbook.SaveAs(WorkbookPath);
        _fileStamp = File.GetLastWriteTimeUtc(WorkbookPath);
        Saved?.Invoke(this, WorkbookPath);
    }

    private static bool TryGetSheet(XLWorkbook workbook, ExcelOp op, out IXLWorksheet sheet)
    {
        if (op.Sheet is not null && workbook.Worksheets.TryGetWorksheet(op.Sheet, out var found) && found is not null)
        {
            sheet = found;
            return true;
        }

        sheet = null!;
        return false;
    }

    /// <summary>Converts the shim's zero-based area into a ClosedXML range; a missing area means "used range".</summary>
    private static bool TryArea(IXLWorksheet sheet, ExcelOp op, out IXLRange range)
    {
        if (op.R is int r && op.C is int c && op.Rows is int rows && op.Cols is int cols && rows > 0 && cols > 0)
        {
            range = sheet.Range(r + 1, c + 1, r + rows, c + cols);
            return true;
        }

        range = null!;
        return false;
    }

    private static IXLTable? ResolveTable(IXLWorksheet sheet, Dictionary<string, IXLTable> created, ExcelOp op)
    {
        if (op.Id is not null && created.TryGetValue(op.Id, out var table)) return table;
        if (op.Target is not null) return sheet.Tables.FirstOrDefault(t => t.Name == op.Target);
        return null;
    }

    private static void ApplyTableProperty(IXLTable? table, ExcelOp op)
    {
        if (table is null || op.Prop is null) return;

        switch (op.Prop)
        {
            case "name":
                var name = op.Value.ValueKind == JsonValueKind.String ? op.Value.GetString() : null;
                if (!string.IsNullOrWhiteSpace(name)) table.Name = name;
                break;
            case "style":
                var style = op.Value.ValueKind == JsonValueKind.String ? op.Value.GetString() : null;
                if (style is not null) table.Theme = ResolveTheme(style);
                break;
            case "showBandedRows":
                table.ShowRowStripes = op.Value.ValueKind != JsonValueKind.False;
                break;
            case "showFilterButton":
                table.SetShowAutoFilter(op.Value.ValueKind != JsonValueKind.False);
                break;
        }
    }

    private static XLTableTheme ResolveTheme(string name)
    {
        var property = typeof(XLTableTheme).GetProperty(name, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
        return property?.GetValue(null) as XLTableTheme ?? XLTableTheme.TableStyleMedium2;
    }

    private static void WriteValues(IXLWorksheet sheet, ExcelOp op)
    {
        if (op.Values.ValueKind != JsonValueKind.Array) return;
        var baseRow = (op.R ?? 0) + 1;
        var baseCol = (op.C ?? 0) + 1;

        var rowIndex = 0;
        foreach (var row in op.Values.EnumerateArray())
        {
            if (row.ValueKind == JsonValueKind.Array)
            {
                var colIndex = 0;
                foreach (var cell in row.EnumerateArray())
                {
                    sheet.Cell(baseRow + rowIndex, baseCol + colIndex).Value = ToCellValue(cell);
                    colIndex++;
                }
            }

            rowIndex++;
        }
    }

    private static void ApplyFormat(IXLWorksheet sheet, ExcelOp op)
    {
        if (op.Props is null || op.Props.Count == 0) return;
        var hasArea = TryArea(sheet, op, out var range);

        foreach (var (prop, value) in op.Props)
        {
            switch (prop)
            {
                case "fontBold":
                    Style(sheet, hasArea, range).Font.Bold = value.ValueKind != JsonValueKind.False;
                    break;
                case "fontItalic":
                    Style(sheet, hasArea, range).Font.Italic = value.ValueKind != JsonValueKind.False;
                    break;
                case "fontColor":
                    if (TryColor(value, out var fontColor)) Style(sheet, hasArea, range).Font.FontColor = fontColor;
                    break;
                case "fontSize":
                    if (value.ValueKind == JsonValueKind.Number) Style(sheet, hasArea, range).Font.FontSize = value.GetDouble();
                    break;
                case "fontName":
                    if (value.ValueKind == JsonValueKind.String && value.GetString() is string fontName)
                        Style(sheet, hasArea, range).Font.FontName = fontName;
                    break;
                case "fillColor":
                    if (TryColor(value, out var fillColor)) Style(sheet, hasArea, range).Fill.BackgroundColor = fillColor;
                    break;
                case "horizontalAlignment":
                    if (TryEnum<XLAlignmentHorizontalValues>(value, out var horizontal))
                        Style(sheet, hasArea, range).Alignment.Horizontal = horizontal;
                    break;
                case "verticalAlignment":
                    if (TryEnum<XLAlignmentVerticalValues>(value, out var vertical))
                        Style(sheet, hasArea, range).Alignment.Vertical = vertical;
                    break;
                case "rowHeight":
                    if (hasArea && value.ValueKind == JsonValueKind.Number)
                    {
                        var height = value.GetDouble();
                        for (var i = 0; i < (op.Rows ?? 0); i++) sheet.Row((op.R ?? 0) + 1 + i).Height = height;
                    }
                    break;
                case "columnWidth":
                    if (hasArea && value.ValueKind == JsonValueKind.Number)
                    {
                        // Office.js uses points; ClosedXML column widths are in character units.
                        var width = value.GetDouble() / 7.0;
                        for (var i = 0; i < (op.Cols ?? 0); i++) sheet.Column((op.C ?? 0) + 1 + i).Width = width;
                    }
                    break;
            }
        }
    }

    private static IXLStyle Style(IXLWorksheet sheet, bool hasArea, IXLRange? range) => hasArea ? range!.Style : sheet.Style;

    private static void ApplyListValidation(IXLRange range, ExcelOp op)
    {
        if (string.IsNullOrEmpty(op.Source)) return;
        // Excel caps an inline validation list at 255 characters.
        if (op.Source.Length > 250) return;

        var validation = range.CreateDataValidation();
        validation.List($"\"{op.Source.Replace("\"", string.Empty)}\"", op.InCellDropDown != false);
    }

    private static object?[][]? ReadValues(IXLWorksheet sheet, ExcelOp op)
    {
        var range = TryArea(sheet, op, out var area) ? area : sheet.RangeUsed();
        if (range is null) return null;

        var rowCount = range.RowCount();
        var colCount = range.ColumnCount();
        var values = new object?[rowCount][];

        for (var r = 0; r < rowCount; r++)
        {
            var row = new object?[colCount];
            for (var c = 0; c < colCount; c++) row[c] = FromCell(range.Cell(r + 1, c + 1));
            values[r] = row;
        }

        return values;
    }

    private static XLCellValue ToCellValue(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.Number => element.GetDouble(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.String => element.GetString() ?? string.Empty,
        JsonValueKind.Null or JsonValueKind.Undefined => Blank.Value,
        _ => element.ToString(),
    };

    private static object? FromCell(IXLCell cell) => cell.Value.Type switch
    {
        XLDataType.Number => cell.Value.GetNumber(),
        XLDataType.Boolean => cell.Value.GetBoolean(),
        XLDataType.DateTime => cell.Value.GetDateTime().ToString("O"),
        XLDataType.TimeSpan => cell.Value.GetTimeSpan().ToString(),
        XLDataType.Text => cell.Value.GetText(),
        XLDataType.Error => string.Empty,
        _ => string.Empty,
    };

    private static string? FirstString(JsonElement element)
    {
        if (element.ValueKind != JsonValueKind.Array) return null;
        foreach (var row in element.EnumerateArray())
        {
            if (row.ValueKind == JsonValueKind.String) return row.GetString();
            if (row.ValueKind != JsonValueKind.Array) continue;
            foreach (var cell in row.EnumerateArray())
            {
                if (cell.ValueKind == JsonValueKind.String) return cell.GetString();
            }
        }

        return null;
    }

    private static bool TryColor(JsonElement element, out XLColor color)
    {
        color = XLColor.NoColor;
        if (element.ValueKind != JsonValueKind.String) return false;
        var text = element.GetString();
        if (string.IsNullOrWhiteSpace(text)) return false;

        try
        {
            color = XLColor.FromHtml(text.StartsWith('#') ? text : "#" + text);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryEnum<T>(JsonElement element, out T value) where T : struct, Enum
    {
        value = default;
        return element.ValueKind == JsonValueKind.String && Enum.TryParse(element.GetString(), true, out value);
    }
}
