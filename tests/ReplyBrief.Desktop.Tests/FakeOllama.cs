using System.Globalization;
using System.Text;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace ReplyBrief.Desktop.Tests;

/// <summary>
/// A loopback HTTP server that speaks the two Ollama routes the desktop app uses. Chat answers are
/// gated on a handle the test releases, so "while the request is in flight" is a state a test can
/// actually hold rather than a race it hopes to win.
/// </summary>
internal sealed class FakeOllama : IAsyncDisposable
{
    private readonly WebApplication _app;
    private readonly TaskCompletionSource _chatGate = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly object _sync = new();
    private readonly bool _holdChat;
    private readonly string? _tagsBody;
    private int _chatRequests;
    private int _disposed;

    private FakeOllama(WebApplication app, bool holdChat, string? tagsBody)
    {
        _app = app;
        _holdChat = holdChat;
        _tagsBody = tagsBody;
    }

    public int Port { get; private set; }

    public string Url => $"http://127.0.0.1:{Port.ToString(CultureInfo.InvariantCulture)}";

    public int ChatRequests
    {
        get
        {
            lock (_sync)
                return _chatRequests;
        }
    }

    /// <summary>Lets every held chat request answer, and every later one answer immediately.</summary>
    public void ReleaseChat() => _chatGate.TrySetResult();

    public static async Task<FakeOllama> StartAsync(bool holdChat = false, string? tagsBody = null)
    {
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            ApplicationName = typeof(FakeOllama).Assembly.GetName().Name,
            ContentRootPath = AppContext.BaseDirectory,
            Args = Array.Empty<string>(),
        });

        builder.Configuration.Sources.Clear();
        builder.Logging.ClearProviders();
        builder.WebHost.ConfigureKestrel(kestrel =>
        {
            kestrel.AddServerHeader = false;
            kestrel.Listen(System.Net.IPAddress.Loopback, 0);
        });

        var server = new FakeOllama(builder.Build(), holdChat, tagsBody);
        server._app.MapPost("/api/chat", server.HandleChatAsync);
        server._app.MapGet("/api/tags", server.HandleTagsAsync);
        await server._app.StartAsync().ConfigureAwait(false);

        foreach (var address in server._app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>()?.Addresses ?? Array.Empty<string>())
        {
            if (!Uri.TryCreate(address, UriKind.Absolute, out var bound) || bound.Port <= 0)
                continue;

            server.Port = bound.Port;
            break;
        }

        return server;
    }

    public async ValueTask DisposeAsync()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
            return;

        _chatGate.TrySetResult();
        await _app.StopAsync().ConfigureAwait(false);
        await _app.DisposeAsync().ConfigureAwait(false);
    }

    private async Task HandleChatAsync(HttpContext context)
    {
        lock (_sync)
            _chatRequests++;

        if (_holdChat)
            await _chatGate.Task.ConfigureAwait(false);

        await WriteAsync(context, StatusCodes.Status200OK, """
            {"done":true,"done_reason":"stop","message":{"role":"assistant","content":"{\"summary\":[\"The deck is due on Thursday.\",\"The discount figure is still open.\"],\"needs_attention\":[\"Thursday is described as tight rather than agreed.\"],\"suggested_reply_points\":[\"Confirm the discount with finance.\"],\"missing_information\":[\"Whether the legal review has landed.\"]}"}}
            """).ConfigureAwait(false);
    }

    private Task HandleTagsAsync(HttpContext context) => WriteAsync(
        context,
        StatusCodes.Status200OK,
        _tagsBody ?? """{"models":[{"name":"qwen2.5:7b"},{"name":"llama3.1"},{"name":"nomic-embed-text"}]}""");

    private static async Task WriteAsync(HttpContext context, int statusCode, string body)
    {
        var bytes = Encoding.UTF8.GetBytes(body);
        context.Response.StatusCode = statusCode;
        context.Response.ContentType = "application/json";
        context.Response.ContentLength = bytes.Length;
        await context.Response.Body.WriteAsync(bytes, context.RequestAborted).ConfigureAwait(false);
    }
}
