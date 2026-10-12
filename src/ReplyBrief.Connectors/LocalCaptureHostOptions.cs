namespace ReplyBrief.Connectors;

/// <summary>
/// Overrides used by the connector test suite. Production callers use
/// <see cref="LocalCaptureHost.StartAsync(System.Action{Core.ConversationCapture}, System.Threading.CancellationToken)"/>,
/// which pins the host to the documented port and asset layout.
/// </summary>
internal sealed record LocalCaptureHostOptions
{
    /// <summary>Zero asks the operating system for a free loopback port, which tests rely on.</summary>
    public int Port { get; init; } = LocalCaptureHost.DefaultPort;

    public string AssetDirectory { get; init; } = AddInAssets.DefaultDirectory;

    /// <summary>
    /// Called when a capture dispatch throws. The host deliberately keeps no logger of its own so
    /// that a capture body or a pairing code can never reach a log sink, which means the desktop
    /// application owns the reporting.
    /// </summary>
    public Action<Exception>? DispatchFailed { get; init; }
}