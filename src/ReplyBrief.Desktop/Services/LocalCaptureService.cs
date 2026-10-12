using ReplyBrief.Connectors;
using ReplyBrief.Core;

namespace ReplyBrief.Desktop.Services;

/// <summary>
/// The only place in the app that names <c>ReplyBrief.Connectors</c>, so the capture host can be
/// swapped without touching the window or the view model.
/// </summary>
internal sealed class LocalCaptureService : IAsyncDisposable
{
    public static readonly string Endpoint = $"https://localhost:{LocalCaptureHost.DefaultPort}";

    private LocalCaptureHost? _host;

    public string PairingCode => _host?.PairingCode ?? string.Empty;

    public static async Task<LocalCaptureService> StartAsync(
        Action<ConversationCapture> onCapture,
        CancellationToken cancellationToken = default)
    {
        var service = new LocalCaptureService();
        service._host = await LocalCaptureHost.StartAsync(onCapture, cancellationToken);
        return service;
    }

    public async ValueTask DisposeAsync()
    {
        var host = _host;
        if (host is null)
            return;

        _host = null;
        await host.DisposeAsync();
    }
}
