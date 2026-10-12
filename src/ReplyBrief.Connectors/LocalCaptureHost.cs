using System.Globalization;
using System.Net;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Channels;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Net.Http.Headers;
using ReplyBrief.Core;

namespace ReplyBrief.Connectors;

/// <summary>
/// A loopback-only HTTPS endpoint that Outlook task pane add-ins post a message to.
/// <para>
/// The host authenticates every caller with a pairing code that only the desktop application knows,
/// revalidates the payload through <see cref="ConversationCapture.Validate"/> and then hands the
/// capture to the callback supplied by the desktop application. What happens next - queueing, model
/// calls, UI - is entirely the desktop's responsibility; the host never writes a payload, the
/// pairing code or any request detail to a log.
/// </para>
/// </summary>
public sealed class LocalCaptureHost : IAsyncDisposable
{
    /// <summary>The fixed loopback port the shipped Outlook manifest points at.</summary>
    public const int DefaultPort = 47831;

    internal const string CapturePath = "/capture";
    internal const string PairingHeaderName = "X-ReplyBrief-Pairing";
    internal const int MaxRequestBytes = 300 * 1024;

    /// <summary>
    /// Captures the desktop app has not picked up yet. The queue is bounded on purpose: an unbounded
    /// one lets a wedged or very slow desktop callback turn repeated captures into unbounded memory
    /// growth. Beyond this the add-in is told to try again rather than being queued silently.
    /// </summary>
    internal const int CaptureQueueCapacity = 8;

    private const string AllowedSource = "outlook";
    private const string CaptureRoute = "/outlook/{asset}";
    private const int DispatchShutdownGraceSeconds = 2;

    private readonly WebApplication _app;
    private readonly Channel<ConversationCapture> _captures;
    private readonly string _pairingCode;
    private readonly string _assetDirectory;
    private readonly Task _dispatch;
    private int _disposed;

    private LocalCaptureHost(
        WebApplication app,
        Channel<ConversationCapture> captures,
        string pairingCode,
        string assetDirectory,
        Action<ConversationCapture> onCapture,
        Action<Exception>? dispatchFailed)
    {
        _app = app;
        _captures = captures;
        _pairingCode = pairingCode;
        _assetDirectory = assetDirectory;

        // The pump runs before the listener opens so a capture can never block on an unstarted
        // consumer, and so a callback that takes its time never delays the task pane response.
        _dispatch = DispatchAsync(captures.Reader, onCapture, dispatchFailed);
    }

    /// <summary>The port Kestrel actually bound, which differs from <see cref="DefaultPort"/> in tests.</summary>
    public int Port { get; private set; } = DefaultPort;

    /// <summary>The single origin the add-in is served from and the only accepted request origin.</summary>
    public Uri BaseAddress => new($"https://localhost:{Port.ToString(CultureInfo.InvariantCulture)}", UriKind.Absolute);

    /// <summary>
    /// The code an operator copies from the desktop application into the task pane. It changes on
    /// every start and is never placed in a URL, an asset, a response header or a log record.
    /// </summary>
    public string PairingCode => _pairingCode;

    /// <summary>Starts the host and returns once it is accepting connections.</summary>
    public static Task<LocalCaptureHost> StartAsync(
        Action<ConversationCapture> onCapture,
        CancellationToken cancellationToken = default) =>
        StartAsync(onCapture, new LocalCaptureHostOptions(), cancellationToken);

    internal static async Task<LocalCaptureHost> StartAsync(
        Action<ConversationCapture> onCapture,
        LocalCaptureHostOptions options,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(onCapture);
        ArgumentNullException.ThrowIfNull(options);

        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            ApplicationName = typeof(LocalCaptureHost).Assembly.GetName().Name,
            ContentRootPath = AppContext.BaseDirectory,
            Args = Array.Empty<string>(),
        });

        // The capture host is entirely self-configured. Clearing the configuration sources keeps an
        // appsettings.json next to the desktop executable, an ASPNETCORE_* environment variable or a
        // stray command line switch from adding a second endpoint, moving the port, or turning on a
        // log sink that could record a request body.
        builder.Configuration.Sources.Clear();
        builder.Logging.ClearProviders();

        builder.WebHost.ConfigureKestrel(kestrel =>
        {
            kestrel.AddServerHeader = false;
            kestrel.Limits.MaxRequestBodySize = MaxRequestBytes;

            // Loopback only, TLS only. UseHttps() with no argument resolves the ASP.NET Core
            // development certificate, so there is no plain HTTP fallback.
            //
            // A fixed port binds both loopback families so "localhost" works whether the resolver
            // hands back 127.0.0.1 or ::1. ListenLocalhost rejects port 0, which is exactly what the
            // tests want, so they fall back to the single-stack IPv4 form.
            if (options.Port == 0)
                kestrel.Listen(IPAddress.Loopback, 0, listen => listen.UseHttps());
            else
                kestrel.ListenLocalhost(options.Port, listen => listen.UseHttps());
        });

        var app = builder.Build();
        var pairingCode = Pairing.Create();
        var captures = Channel.CreateBounded<ConversationCapture>(new BoundedChannelOptions(CaptureQueueCapacity)
        {
            SingleReader = true,
            FullMode = BoundedChannelFullMode.Wait,
        });
        var host = new LocalCaptureHost(app, captures, pairingCode, options.AssetDirectory, onCapture, options.DispatchFailed);

        try
        {
            app.MapPost(CapturePath, host.HandleCaptureAsync);
            app.MapGet(CaptureRoute, host.ServeAddInAssetAsync);
            await app.StartAsync(cancellationToken).ConfigureAwait(false);
        }
        catch
        {
            captures.Writer.TryComplete();
            await host._dispatch.ConfigureAwait(false);
            await app.DisposeAsync().ConfigureAwait(false);
            throw;
        }

        host.Port = ResolvePort(app, options.Port);
        return host;
    }

    public async ValueTask DisposeAsync()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;

        _captures.Writer.TryComplete();

        // Let an in-flight capture finish handing off; never hang shutdown on a slow callback.
        try
        {
            await _dispatch.WaitAsync(TimeSpan.FromSeconds(DispatchShutdownGraceSeconds)).ConfigureAwait(false);
        }
        catch (TimeoutException)
        {
        }

        await _app.StopAsync().ConfigureAwait(false);
        await _app.DisposeAsync().ConfigureAwait(false);
    }

    private async Task HandleCaptureAsync(HttpContext context)
    {
        if (ScreenRequest(context) is { } rejection)
        {
            await RejectAsync(context, rejection.StatusCode, rejection.Reason).ConfigureAwait(false);
            return;
        }

        CapturePayload? payload;
        try
        {
            payload = await ReadPayloadAsync(context.Request).ConfigureAwait(false);
        }
        catch (JsonException)
        {
            await RejectAsync(context, StatusCodes.Status400BadRequest, "The capture was not well-formed JSON.").ConfigureAwait(false);
            return;
        }
        catch (BadHttpRequestException unreadable)
        {
            // Kestrel raises this for a body over the limit, a truncated chunked body or a header
            // block it will not accept. Its own StatusCode is the right answer (413, 400, 408...),
            // but its message can quote the request, so a fixed reason is returned instead.
            await RejectAsync(context, unreadable.StatusCode, "The request could not be read within the capture limits.").ConfigureAwait(false);
            return;
        }

        var source = payload?.Source;
        if (payload is null)
        {
            await RejectAsync(context, StatusCodes.Status400BadRequest, "The capture must be a JSON object.").ConfigureAwait(false);
            return;
        }

        if (source is not AllowedSource)
        {
            await RejectAsync(context, StatusCodes.Status400BadRequest, "Only captures from the Outlook add-in are accepted.").ConfigureAwait(false);
            return;
        }

        var capture = new ConversationCapture(
            Source: source,
            Title: payload.Title ?? "",
            Participants: payload.Participants ?? "",
            Content: payload.Content ?? "",
            ContextUsed: payload.ContextUsed ?? "");

        try
        {
            // Core owns the field limits and the accepted source vocabulary; reusing its validation
            // keeps the wire contract and the rest of the application from drifting apart.
            capture.Validate();
        }
        catch (ArgumentException invalid)
        {
            await RejectAsync(context, StatusCodes.Status400BadRequest, invalid.Message).ConfigureAwait(false);
            return;
        }

        // The desktop app owns the durable queue; this host only refuses to accept work it cannot hand
        // over. A disposed host is shutting down; a full queue means the desktop is behind.
        if (Volatile.Read(ref _disposed) != 0)
        {
            await RejectAsync(context, StatusCodes.Status503ServiceUnavailable, "ReplyBrief is shutting down and is not accepting new captures.").ConfigureAwait(false);
            return;
        }

        if (!_captures.Writer.TryWrite(capture))
        {
            await RejectAsync(context, Volatile.Read(ref _disposed) != 0 ? StatusCodes.Status503ServiceUnavailable : StatusCodes.Status429TooManyRequests,
                Volatile.Read(ref _disposed) != 0
                    ? "ReplyBrief is shutting down and is not accepting new captures."
                    : $"ReplyBrief already has {CaptureQueueCapacity} captures waiting. Try again in a moment.").ConfigureAwait(false);
            return;
        }

        await WriteAsync(context, StatusCodes.Status202Accepted, new CaptureResponse("accepted")).ConfigureAwait(false);
    }

    /// <summary>
    /// Everything a request has to clear before its body is even looked at, in the order that leaks
    /// the least: whether the host is still running, where it came from, who it is, how it is encoded
    /// and how big it is allowed to be.
    /// </summary>
    private (int StatusCode, string Reason)? ScreenRequest(HttpContext context)
    {
        if (Volatile.Read(ref _disposed) != 0)
            return (StatusCodes.Status503ServiceUnavailable, "ReplyBrief is shutting down and is not accepting new captures.");

        if (!IsSameOrigin(context))
            return (StatusCodes.Status403Forbidden, "This request did not originate from the ReplyBrief task pane.");

        var presented = context.Request.Headers[PairingHeaderName];
        if (presented.Count != 1 || !Pairing.Matches(_pairingCode, presented[0]))
            return (StatusCodes.Status401Unauthorized, "The pairing code is missing or no longer current.");

        if (!IsJsonContentType(context.Request))
            return (StatusCodes.Status415UnsupportedMediaType, "The capture must be sent as application/json.");

        if (context.Request.ContentLength > MaxRequestBytes)
            return (StatusCodes.Status413PayloadTooLarge, "The capture is larger than the 300 KB limit.");

        return null;
    }

    private async Task ServeAddInAssetAsync(HttpContext context, string asset)
    {
        if (!AddInAssets.TryResolveFileName($"/{AddInAssets.DirectoryName}/{asset}", out var fileName))
        {
            await RejectAsync(context, StatusCodes.Status404NotFound, "Unknown asset.").ConfigureAwait(false);
            return;
        }

        var path = Path.Combine(_assetDirectory, fileName);
        if (!File.Exists(path))
        {
            await RejectAsync(context, StatusCodes.Status404NotFound, "The Outlook add-in assets are missing from the ReplyBrief installation.").ConfigureAwait(false);
            return;
        }

        var bytes = await File.ReadAllBytesAsync(path, context.RequestAborted).ConfigureAwait(false);

        context.Response.StatusCode = StatusCodes.Status200OK;
        context.Response.ContentType = AddInAssets.ContentTypeFor(fileName);
        context.Response.ContentLength = bytes.Length;
        context.Response.Headers.CacheControl = "no-store";
        context.Response.Headers.XContentTypeOptions = "nosniff";
        context.Response.Headers.ContentSecurityPolicy = AddInAssets.ContentSecurityPolicy;
        await context.Response.Body.WriteAsync(bytes, context.RequestAborted).ConfigureAwait(false);
    }

    /// <summary>
    /// Browsers attach <c>Origin</c> to cross-origin requests and to every non-GET request, so a
    /// task pane posting from the host's own origin is the only case that survives this check.
    /// The comparison is against the port the connection actually landed on rather than the
    /// client-supplied <c>Host</c> header, which an attacker controls.
    /// </summary>
    private static bool IsSameOrigin(HttpContext context)
    {
        var origin = context.Request.Headers.Origin.ToString();
        if (string.IsNullOrEmpty(origin)) return true;

        if (!Uri.TryCreate(origin, UriKind.Absolute, out var uri)) return false;
        if (!string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) return false;

        var port = context.Connection.LocalPort;
        if (port <= 0) return false;

        return string.Equals(uri.Authority, $"localhost:{port.ToString(CultureInfo.InvariantCulture)}", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsJsonContentType(HttpRequest request) =>
        !string.IsNullOrWhiteSpace(request.ContentType)
        && MediaTypeHeaderValue.TryParse(request.ContentType, out var parsed)
        && string.Equals(parsed.MediaType.Value, "application/json", StringComparison.OrdinalIgnoreCase);

    private static async Task<CapturePayload?> ReadPayloadAsync(HttpRequest request)
    {
        using var document = await JsonDocument.ParseAsync(
            request.Body,
            new JsonDocumentOptions { MaxDepth = 8, AllowTrailingCommas = false, CommentHandling = JsonCommentHandling.Disallow },
            request.HttpContext.RequestAborted).ConfigureAwait(false);

        if (document.RootElement.ValueKind != JsonValueKind.Object) return null;

        return new CapturePayload(
            ReadString(document.RootElement, "source"),
            ReadString(document.RootElement, "title"),
            ReadString(document.RootElement, "participants"),
            ReadString(document.RootElement, "content"),
            ReadString(document.RootElement, "contextUsed"));

        static string? ReadString(JsonElement root, string name) =>
            root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
    }

    private static Task WriteAsync(HttpContext context, int statusCode, CaptureResponse body)
    {
        context.Response.StatusCode = statusCode;
        context.Response.Headers.CacheControl = "no-store";
        context.Response.Headers.XContentTypeOptions = "nosniff";
        return context.Response.WriteAsJsonAsync(body, context.RequestAborted);
    }

    private static Task RejectAsync(HttpContext context, int statusCode, string reason) =>
        WriteAsync(context, statusCode, new CaptureResponse("rejected", reason));

    private static async Task DispatchAsync(
        ChannelReader<ConversationCapture> reader,
        Action<ConversationCapture> onCapture,
        Action<Exception>? onFailure)
    {
        await foreach (var capture in reader.ReadAllAsync().ConfigureAwait(false))
        {
            try
            {
                onCapture(capture);
            }
            catch (Exception failure)
            {
                try
                {
                    onFailure?.Invoke(failure);
                }
                catch (Exception)
                {
                    // A reporting callback must not be able to stop the pump; the next capture still
                    // has to be delivered.
                }
            }
        }
    }

    private static int ResolvePort(WebApplication app, int fallback)
    {
        var addresses = app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>()?.Addresses;
        foreach (var address in addresses ?? Array.Empty<string>())
        {
            if (Uri.TryCreate(address, UriKind.Absolute, out var uri) && uri.Port > 0) return uri.Port;
        }

        return fallback;
    }

    private sealed record CaptureResponse(
        string Status,
        [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Error = null);

    private sealed record CapturePayload(string? Source, string? Title, string? Participants, string? Content, string? ContextUsed);
}