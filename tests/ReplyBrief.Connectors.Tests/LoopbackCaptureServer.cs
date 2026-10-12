using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using ReplyBrief.Core;

namespace ReplyBrief.Connectors.Tests;

/// <summary>
/// Starts the real capture host on an operating system assigned loopback port and gives tests an
/// HTTPS client for it. Nothing here relaxes the machine's trust configuration: the client accepts
/// exactly the certificate that the ASP.NET development certificate provider installed in the
/// current user's personal store, and nothing else.
/// </summary>
internal sealed class LoopbackCaptureServer : IAsyncDisposable
{
    private readonly LocalCaptureHost _host;
    private readonly List<Exception> _dispatchFailures;
    private readonly HttpClient _client;

    private LoopbackCaptureServer(LocalCaptureHost host, List<Exception> dispatchFailures)
    {
        _host = host;
        _dispatchFailures = dispatchFailures;
        _client = new HttpClient(new HttpClientHandler { ServerCertificateCustomValidationCallback = AcceptDevelopmentCertificate })
        {
            Timeout = TimeSpan.FromSeconds(30),
        };
    }

    public Uri BaseAddress => _host.BaseAddress;

    public int Port => _host.Port;

    public string PairingCode => _host.PairingCode;

    public IReadOnlyList<Exception> DispatchFailures
    {
        get
        {
            lock (_dispatchFailures) return _dispatchFailures.ToArray();
        }
    }

    public static async Task<LoopbackCaptureServer> StartAsync(Action<ConversationCapture> onCapture, int port = 0)
    {
        var dispatchFailures = new List<Exception>();
        var host = await LocalCaptureHost.StartAsync(onCapture, new LocalCaptureHostOptions
        {
            Port = port,
            DispatchFailed = failure => { lock (dispatchFailures) dispatchFailures.Add(failure); },
        });

        return new LoopbackCaptureServer(host, dispatchFailures);
    }

    /// <summary>
    /// Claims a port that is free on both loopback families, so a host bound with ListenLocalhost can
    /// bind it and the test can reach it over IPv4 and IPv6.
    /// </summary>
    public static int FreeLoopbackPort()
    {
        var listener = new System.Net.Sockets.TcpListener(IPAddress.IPv6Any, 0);
        listener.Start();

        try
        {
            return ((IPEndPoint)listener.LocalEndpoint).Port;
        }
        finally
        {
            listener.Stop();
        }
    }

    /// <summary>
    /// Issues a request over a specific loopback family. The dev certificate is accepted only for
    /// loopback, so this never weakens anything outside the machine.
    /// </summary>
    public static async Task<HttpResponseMessage> SendToAsync(Uri origin, HttpRequestMessage request)
    {
        using var client = new HttpClient(new HttpClientHandler { ServerCertificateCustomValidationCallback = AcceptDevelopmentCertificate })
        {
            Timeout = TimeSpan.FromSeconds(30),
        };

        return await client.SendAsync(request);
    }

    /// <summary>Builds a signed capture POST aimed at this server.</summary>
    public HttpRequestMessage CaptureRequest(string body)
    {
        var request = new HttpRequestMessage(HttpMethod.Post, new Uri(BaseAddress, LocalCaptureHost.CapturePath))
        {
            Content = new ByteArrayContent(Encoding.UTF8.GetBytes(body)),
        };
        request.Content.Headers.TryAddWithoutValidation("Content-Type", "application/json");
        request.Headers.TryAddWithoutValidation(LocalCaptureHost.PairingHeaderName, PairingCode);
        request.Headers.TryAddWithoutValidation("Origin", BaseAddress.GetLeftPart(UriPartial.Authority));
        return request;
    }

    public Task<HttpResponseMessage> PostCaptureAsync(
        string body,
        string? pairingCode,
        string? contentType = "application/json",
        string? origin = null)
    {
        var content = new ByteArrayContent(Encoding.UTF8.GetBytes(body));
        if (contentType is not null) content.Headers.TryAddWithoutValidation("Content-Type", contentType);

        var request = new HttpRequestMessage(HttpMethod.Post, new Uri(BaseAddress, LocalCaptureHost.CapturePath)) { Content = content };
        if (pairingCode is not null) request.Headers.TryAddWithoutValidation(LocalCaptureHost.PairingHeaderName, pairingCode);
        if (origin is not null) request.Headers.TryAddWithoutValidation("Origin", origin);

        return _client.SendAsync(request);
    }

    public Task<HttpResponseMessage> GetAsync(string relativePath) =>
        _client.GetAsync(new Uri(BaseAddress, relativePath));

    public async ValueTask DisposeAsync()
    {
        _client.Dispose();
        await _host.DisposeAsync();
    }

    private static bool AcceptDevelopmentCertificate(HttpRequestMessage request, X509Certificate2? certificate, X509Chain? chain, SslPolicyErrors errors)
    {
        if (certificate is null) return false;
        if (!string.Equals(request.RequestUri?.Scheme, Uri.UriSchemeHttps, StringComparison.Ordinal)) return false;
        if (request.RequestUri?.Host is not ("localhost" or "127.0.0.1" or "::1" or "[::1]")) return false;

        // The test neither disables certificate validation nor writes to the trust store. It accepts
        // only the leaf certificate that is installed in the current user's personal store, presented
        // over loopback, which is exactly the configuration the host is documented to use.
        using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);
        return store.Certificates.Find(X509FindType.FindByThumbprint, certificate.Thumbprint, validOnly: false).Count > 0;
    }
}

/// <summary>Collects dispatched captures so a test can await the one it cares about.</summary>
internal sealed class CaptureRecorder
{
    private readonly TaskCompletionSource<ConversationCapture> _first = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private int _count;

    public int Count => Volatile.Read(ref _count);

    public Task<ConversationCapture> NextAsync() => _first.Task.WaitAsync(TimeSpan.FromSeconds(15));

    public void Record(ConversationCapture capture)
    {
        Interlocked.Increment(ref _count);
        _first.TrySetResult(capture);
    }

    public static string OutlookCaptureJson(string? content = "Hello, are we still on for Thursday?", string source = "outlook") =>
        $$"""
        {
          "source": {{JsonSerializer.Serialize(source)}},
          "title": "Project sync",
          "participants": "Dana Example <dana@example.com>",
          "content": {{JsonSerializer.Serialize(content)}},
          "contextUsed": "Outlook reading pane"
        }
        """;
}