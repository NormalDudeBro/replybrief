namespace ReplyBrief.Connectors;

/// <summary>
/// The complete, closed list of files the capture host is willing to hand to a browser. Serving is
/// table driven rather than directory driven: a request is resolved by exact route match against
/// this table, so <c>%2e%2e</c>, absolute paths and new files dropped into the output directory are
/// unreachable by construction, not by filtering.
/// </summary>
internal static class AddInAssets
{
    internal const string DirectoryName = "outlook";

    /// <summary>
    /// Locked down to what the add-in needs. office.js genuinely calls out to the Office service
    /// endpoints while it initialises, which is why the Microsoft and Office domains appear in
    /// <c>connect-src</c>.
    /// <para>
    /// <c>frame-ancestors</c> and <c>frame-src</c> list the Office and Outlook origins broadly on
    /// purpose. Office embeds a task pane from several different hosts and tenant domains, and the
    /// runtime can itself open a nested frame; a narrower list blanks the pane with no useful error
    /// rather than degrading. ReplyBrief frames nothing itself and there is no content-injection sink
    /// for these to abuse, so the allowance is narrow in origin count even though it is broad in
    /// shape. It is deliberately not a blanket <c>*</c>.
    /// </para>
    /// </summary>
    internal const string ContentSecurityPolicy =
        "default-src 'none'; " +
        "script-src 'self' https://appsforoffice.microsoft.com; " +
        "style-src 'self' 'unsafe-inline'; " +
        "img-src 'self' data:; " +
        "font-src 'self'; " +
        "connect-src 'self' https://*.microsoft.com https://*.office.com https://*.office365.com; " +
        "frame-src 'self' https://*.office.com https://*.office365.com; " +
        "base-uri 'none'; " +
        "form-action 'none'; " +
        "object-src 'none'; " +
        "frame-ancestors 'self' " +
        "https://outlook.com https://*.outlook.com " +
        "https://outlook.office.com https://outlook.office365.com " +
        "https://*.office.com https://*.office365.com " +
        "https://outlook.live.com https://*.live.com";

    private static readonly KeyValuePair<string, string>[] Files =
    [
        Key("/outlook/taskpane.html", "taskpane.html"),
        Key("/outlook/taskpane.js", "taskpane.js"),
        Key("/outlook/commands.html", "commands.html"),
        Key("/outlook/support.html", "support.html"),
        Key("/outlook/icon.svg", "icon.svg"),
        Key("/outlook/icon-16.png", "icon-16.png"),
        Key("/outlook/icon-32.png", "icon-32.png"),
        Key("/outlook/icon-80.png", "icon-80.png"),
        Key("/outlook/icon-128.png", "icon-128.png"),
        Key("/outlook/manifest.xml", "manifest.xml"),
    ];

    private static readonly Dictionary<string, string> ByRoute =
        new(Files, StringComparer.OrdinalIgnoreCase);

    internal static string DefaultDirectory => Path.Combine(AppContext.BaseDirectory, DirectoryName);

    internal static bool TryResolveFileName(string route, out string fileName) =>
        ByRoute.TryGetValue(route, out fileName!);

    internal static string ContentTypeFor(string fileName) =>
        Path.GetExtension(fileName).ToLowerInvariant() switch
        {
            ".html" => "text/html; charset=utf-8",
            ".js" => "text/javascript; charset=utf-8",
            ".svg" => "image/svg+xml",
            ".png" => "image/png",
            ".xml" => "application/xml; charset=utf-8",
            _ => "application/octet-stream",
        };

    private static KeyValuePair<string, string> Key(string route, string fileName) =>
        KeyValuePair.Create(route, fileName);
}