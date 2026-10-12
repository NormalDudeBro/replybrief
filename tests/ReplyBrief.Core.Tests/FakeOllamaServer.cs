using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace ReplyBrief.Core.Tests;

/// <summary>One chat request exactly as the fake Ollama API received it on the wire.</summary>
internal sealed record ChatRequest(string Method, string Path, string ContentType, string Body, JsonElement Root);

/// <summary>
/// A loopback-only HTTP server that speaks the slice of the Ollama chat protocol
/// <see cref="ConversationAnalyzer"/> uses. It records the request it was given and then writes back
/// exactly what the test scripted: there is no model behind it, so every assertion about what the
/// analyzer understands is an assertion about the analyzer and never about this fixture.
/// </summary>
internal sealed class FakeOllamaServer : IAsyncDisposable
{
    /// <summary>The Ollama route the analyzer posts to under a bare server address.</summary>
    public const string ChatRoute = "/api/chat";

    private readonly WebApplication _app;
    private readonly Func<HttpRequest, Task> _responder;
    private readonly List<ChatRequest> _requests = new();
    private int _disposed;

    private FakeOllamaServer(WebApplication app, Func<HttpRequest, Task> responder)
    {
        _app = app;
        _responder = responder;
    }

    public int Port { get; private set; }

    /// <summary>The server as an IPv4 loopback literal.</summary>
    public string Url => $"http://127.0.0.1:{Port.ToString(CultureInfo.InvariantCulture)}";

    /// <summary>The same server reached by name, so a test also exercises name resolution.</summary>
    public string LocalhostUrl => $"http://localhost:{Port.ToString(CultureInfo.InvariantCulture)}";

    public IReadOnlyList<ChatRequest> Requests
    {
        get
        {
            lock (_requests) return _requests.ToArray();
        }
    }

    /// <summary>Answers every request with the same scripted body.</summary>
    public static Task<FakeOllamaServer> StartAsync(
        string body,
        int statusCode = StatusCodes.Status200OK,
        string route = ChatRoute) =>
        StartAsync(request => FakeOllama.WriteAsync(request, statusCode, body), route);

    /// <summary>Answers each request with whatever the test's responder writes.</summary>
    public static async Task<FakeOllamaServer> StartAsync(
        Func<HttpRequest, Task> responder,
        string route = ChatRoute)
    {
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            ApplicationName = typeof(FakeOllamaServer).Assembly.GetName().Name,
            ContentRootPath = AppContext.BaseDirectory,
            Args = Array.Empty<string>(),
        });

        // The fixture answers the requests a test scripts and nothing else. Clearing the sources keeps
        // an appsettings.json, an ASPNETCORE_* variable or a stray switch from adding a second
        // endpoint or starting a log sink that could print a request body.
        builder.Configuration.Sources.Clear();
        builder.Logging.ClearProviders();

        builder.WebHost.ConfigureKestrel(kestrel =>
        {
            kestrel.AddServerHeader = false;
            kestrel.Listen(IPAddress.Loopback, 0);
        });

        var server = new FakeOllamaServer(builder.Build(), responder);
        server._app.MapPost(route, server.HandleAsync);
        server._app.MapGet(route, server.HandleAsync);
        await server._app.StartAsync().ConfigureAwait(false);

        foreach (var address in server._app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>()?.Addresses ?? Array.Empty<string>())
        {
            if (!Uri.TryCreate(address, UriKind.Absolute, out var bound) || bound.Port <= 0) continue;
            server.Port = bound.Port;
            break;
        }

        return server;
    }

    public async ValueTask DisposeAsync()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        await _app.StopAsync().ConfigureAwait(false);
        await _app.DisposeAsync().ConfigureAwait(false);
    }

    private async Task HandleAsync(HttpContext context)
    {
        var body = await ReadBodyAsync(context.Request).ConfigureAwait(false);
        lock (_requests)
        {
            _requests.Add(new ChatRequest(
                context.Request.Method,
                context.Request.Path.Value ?? string.Empty,
                context.Request.ContentType ?? string.Empty,
                body.Text,
                body.Root));
        }

        await _responder(context.Request).ConfigureAwait(false);
    }

    private static async Task<(string Text, JsonElement Root)> ReadBodyAsync(HttpRequest request)
    {
        using var buffer = new MemoryStream();
        await request.Body.CopyToAsync(buffer, request.HttpContext.RequestAborted).ConfigureAwait(false);
        var text = Encoding.UTF8.GetString(buffer.ToArray());
        if (string.IsNullOrWhiteSpace(text)) return (text, default);

        using var document = JsonDocument.Parse(text);
        return (text, document.RootElement.Clone());
    }
}