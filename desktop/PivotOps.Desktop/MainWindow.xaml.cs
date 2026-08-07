using System.Text.Json;
using Microsoft.UI.Xaml;
using Microsoft.Web.WebView2.Core;
using PivotOps.Desktop.Services;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.System;

namespace PivotOps.Desktop;

public sealed partial class MainWindow : Window
{
    private const string VirtualHost = "pivotops.local";
    private const string StartPage = $"https://{VirtualHost}/taskpane.html";

    private readonly WorkbookHost _workbook;
    private readonly SecureStore _store = new();

    public MainWindow()
    {
        InitializeComponent();

        Title = "PivotOps";
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(TitleBar);
        AppWindow.Resize(new Windows.Graphics.SizeInt32(1180, 860));

        _workbook = new WorkbookHost(Path.Combine(GetDataFolder(), "PivotOps.xlsx"));
        _workbook.Saved += OnWorkbookSaved;

        UpdateWorkbookButtons();
        _ = InitializeWebViewAsync();
    }

    private static string GetDataFolder()
    {
        try
        {
            return ApplicationData.Current.LocalFolder.Path;
        }
        catch
        {
            var fallback = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PivotOps");
            Directory.CreateDirectory(fallback);
            return fallback;
        }
    }

    private async Task InitializeWebViewAsync()
    {
        try
        {
            await WebView.EnsureCoreWebView2Async();
        }
        catch (Exception ex)
        {
            ShowFatal(
                "The Microsoft Edge WebView2 Runtime is required to run PivotOps.\n\n" +
                "Install it from https://developer.microsoft.com/microsoft-edge/webview2/ and restart the app.\n\n" +
                ex.Message);
            return;
        }

        var core = WebView.CoreWebView2;
        var webRoot = Path.Combine(AppContext.BaseDirectory, "web");
        if (!File.Exists(Path.Combine(webRoot, "taskpane.html")))
        {
            ShowFatal($"The bundled web app was not found at:\n{webRoot}\n\nRun `npm run build:desktop` and rebuild.");
            return;
        }

        core.SetVirtualHostNameToFolderMapping(VirtualHost, webRoot, CoreWebView2HostResourceAccessKind.DenyCors);
        core.Settings.AreDevToolsEnabled = false;
        core.Settings.IsStatusBarEnabled = false;
        core.Settings.IsSwipeNavigationEnabled = false;
        core.Settings.AreBrowserAcceleratorKeysEnabled = false;

        core.WebMessageReceived += OnWebMessageReceived;
        core.NewWindowRequested += OnNewWindowRequested;

        core.Navigate(StartPage);
    }

    private static async void OnNewWindowRequested(CoreWebView2 sender, CoreWebView2NewWindowRequestedEventArgs args)
    {
        args.Handled = true;
        if (Uri.TryCreate(args.Uri, UriKind.Absolute, out var uri) && (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp))
        {
            await Launcher.LaunchUriAsync(uri);
        }
    }

    // ── Bridge ─────────────────────────────────────────────────────────────

    private async void OnWebMessageReceived(CoreWebView2 sender, CoreWebView2WebMessageReceivedEventArgs args)
    {
        HostRequest? request;
        try
        {
            request = JsonSerializer.Deserialize<HostRequest>(args.TryGetWebMessageAsString(), BridgeJson.Options);
        }
        catch
        {
            return;
        }

        if (request is null) return;

        try
        {
            Respond(new HostResponse { Id = request.Id, Ok = true, Result = await DispatchAsync(request) });
        }
        catch (Exception ex)
        {
            Respond(new HostResponse { Id = request.Id, Ok = false, Error = ex.Message });
        }
    }

    private async Task<object?> DispatchAsync(HostRequest request)
    {
        switch (request.Kind)
        {
            case "excel.sync":
            {
                var payload = request.Payload.Deserialize<ExcelSyncPayload>(BridgeJson.Options) ?? new ExcelSyncPayload();
                var results = await _workbook.ApplyAsync(payload.Ops);
                return new { results };
            }

            case "storage.get":
                return _store.Get(RequireKey(request));

            case "storage.set":
            {
                var payload = request.Payload.Deserialize<StoragePayload>(BridgeJson.Options)!;
                _store.Set(payload.Key, payload.Value ?? string.Empty);
                return null;
            }

            case "storage.remove":
                _store.Remove(RequireKey(request));
                return null;

            default:
                throw new NotSupportedException($"Unknown host request '{request.Kind}'.");
        }
    }

    private static string RequireKey(HostRequest request)
    {
        var payload = request.Payload.Deserialize<StoragePayload>(BridgeJson.Options);
        if (payload is null || string.IsNullOrEmpty(payload.Key)) throw new ArgumentException("Missing storage key.");
        return payload.Key;
    }

    private void Respond(HostResponse response)
    {
        var json = JsonSerializer.Serialize(response, BridgeJson.Options);
        DispatcherQueue.TryEnqueue(() => WebView.CoreWebView2?.PostWebMessageAsString(json));
    }

    // ── Workbook commands ──────────────────────────────────────────────────

    private void OnWorkbookSaved(object? sender, string path)
    {
        DispatcherQueue.TryEnqueue(() =>
        {
            StatusText.Text = path;
            UpdateWorkbookButtons();
        });
    }

    private void UpdateWorkbookButtons()
    {
        var exists = _workbook.Exists;
        OpenInExcelButton.IsEnabled = exists;
        SaveCopyButton.IsEnabled = exists;
        RevealButton.IsEnabled = exists;
        if (exists) StatusText.Text = _workbook.WorkbookPath;
    }

    private async void OnOpenWorkbook(object sender, RoutedEventArgs e)
    {
        if (!_workbook.Exists) return;
        var file = await StorageFile.GetFileFromPathAsync(_workbook.WorkbookPath);
        await Launcher.LaunchFileAsync(file);
    }

    private async void OnRevealWorkbook(object sender, RoutedEventArgs e)
    {
        if (!_workbook.Exists) return;
        var file = await StorageFile.GetFileFromPathAsync(_workbook.WorkbookPath);
        var folder = await file.GetParentAsync();
        var options = new FolderLauncherOptions();
        options.ItemsToSelect.Add(file);
        await Launcher.LaunchFolderAsync(folder, options);
    }

    private async void OnSaveCopy(object sender, RoutedEventArgs e)
    {
        if (!_workbook.Exists) return;

        var picker = new FileSavePicker { SuggestedStartLocation = PickerLocationId.DocumentsLibrary };
        picker.FileTypeChoices.Add("Excel workbook", new List<string> { ".xlsx" });
        picker.SuggestedFileName = "PivotOps";
        WinRT.Interop.InitializeWithWindow.Initialize(picker, WinRT.Interop.WindowNative.GetWindowHandle(this));

        var target = await picker.PickSaveFileAsync();
        if (target is null) return;

        var source = await StorageFile.GetFileFromPathAsync(_workbook.WorkbookPath);
        await source.CopyAndReplaceAsync(target);
    }

    private void ShowFatal(string message)
    {
        LoadError.Text = message;
        LoadError.Visibility = Visibility.Visible;
        WebView.Visibility = Visibility.Collapsed;
    }
}
