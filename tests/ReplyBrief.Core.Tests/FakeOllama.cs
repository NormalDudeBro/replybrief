using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Http;

namespace ReplyBrief.Core.Tests;

/// <summary>
/// The canned Ollama replies and response framings the tests need. Everything here is protocol, not
/// model behaviour: each helper writes a body that is either exactly what the analyzer asked for, or
/// one specific way of failing to give it that.
/// </summary>
internal static class FakeOllama
{
    /// <summary>A complete analysis, exactly the four sections the requested schema describes.</summary>
    public const string AnalysisContent =
        """
        {"summary":["Dana asked whether the launch date moved to Thursday."],"needs_attention":["Dana wants an answer before Tuesday's review."],"suggested_reply_points":["Confirm whether the Thursday date still holds."],"missing_information":["Whether the vendor has confirmed the new date."]}
        """;

    /// <summary>Wraps model content in the chat envelope the analyzer reads.</summary>
    public static string Chat(string content, string doneReason = "stop") =>
        """{"model":"replybrief-test","created_at":"2026-01-01T00:00:00Z","done":true,"done_reason":"""
        + JsonSerializer.Serialize(doneReason)
        + ""","total_duration":1,"prompt_eval_count":1,"eval_count":1,"message":{"role":"assistant","content":"""
        + JsonSerializer.Serialize(content)
        + "}}";

    /// <summary>The envelope a healthy model returns for the happy path.</summary>
    public static string Analysis() => Chat(AnalysisContent);

    /// <summary>
    /// A well-formed analysis carrying a field the analyzer ignores, so a body can be made large
    /// enough to need more than one read without changing what the analyzer should accept.
    /// </summary>
    public static string PaddedAnalysis(int paddingCharacters) =>
        """{"model":"replybrief-test","done":true,"done_reason":"stop","raw_padding":"""
        + JsonSerializer.Serialize(new string('p', paddingCharacters))
        + ""","message":{"role":"assistant","content":"""
        + JsonSerializer.Serialize(AnalysisContent)
        + "}}";

    /// <summary>A response the analyzer has to refuse on size, whatever its content says.</summary>
    public static string OversizedAnalysis(int minimumCharacters = 256 * 1024) => Chat(new string('x', minimumCharacters));

    /// <summary>Builds an analysis body from named sections, for the cases that are deliberately wrong.</summary>
    public static string Content(params (string Name, string[] Bullets)[] sections) =>
        JsonSerializer.Serialize(sections.ToDictionary(section => section.Name, section => section.Bullets));

    /// <summary>Writes a complete response with a declared length.</summary>
    public static Task WriteAsync(HttpRequest request, int statusCode, string body)
    {
        var context = request.HttpContext;
        var bytes = Encoding.UTF8.GetBytes(body);
        context.Response.StatusCode = statusCode;
        context.Response.ContentType = "application/json";
        context.Response.ContentLength = bytes.Length;
        return context.Response.Body.WriteAsync(bytes, 0, bytes.Length, context.RequestAborted);
    }

    /// <summary>Answers with a redirect, which is what an Ollama route behind a proxy can do.</summary>
    public static Task RedirectAsync(HttpRequest request, string location)
    {
        var response = request.HttpContext.Response;
        response.StatusCode = StatusCodes.Status302Found;
        response.Headers.Location = location;
        response.ContentLength = 0;
        return Task.CompletedTask;
    }

    /// <summary>
    /// Writes the body in pieces with no declared length, so the analyzer has to read it as a stream.
    /// Flushing between pieces is what commits Kestrel to chunked framing.
    /// </summary>
    public static async Task WriteChunkedAsync(HttpRequest request, string body, int chunkSize = 16 * 1024)
    {
        var context = request.HttpContext;
        context.Response.StatusCode = StatusCodes.Status200OK;
        context.Response.ContentType = "application/json";

        var bytes = Encoding.UTF8.GetBytes(body);
        for (var offset = 0; offset < bytes.Length; offset += chunkSize)
        {
            // The analyzer stops reading the moment a body is too large, so the rest of the response
            // may never be wanted.
            if (context.RequestAborted.IsCancellationRequested) return;
            var count = Math.Min(chunkSize, bytes.Length - offset);
            await context.Response.Body.WriteAsync(bytes, offset, count, context.RequestAborted).ConfigureAwait(false);
            await context.Response.Body.FlushAsync(context.RequestAborted).ConfigureAwait(false);
        }
    }

    /// <summary>Sends headers and the first slice of the body, then leaves the response open.</summary>
    public static async Task WritePartialAsync(HttpRequest request)
    {
        var context = request.HttpContext;
        context.Response.StatusCode = StatusCodes.Status200OK;
        context.Response.ContentType = "application/json";

        var prefix = Encoding.UTF8.GetBytes("""{"done":true,"done_reason":"stop","message":{"role":"assistant","content":""");
        await context.Response.Body.WriteAsync(prefix, 0, prefix.Length, context.RequestAborted).ConfigureAwait(false);
        await context.Response.Body.FlushAsync(context.RequestAborted).ConfigureAwait(false);
    }

    /// <summary>Answers nothing at all until the test releases the request.</summary>
    public static async Task HoldAsync(HttpRequest request, Task release)
    {
        await release.ConfigureAwait(false);
        if (request.HttpContext.RequestAborted.IsCancellationRequested) return;
        await WriteAsync(request, StatusCodes.Status200OK, Analysis()).ConfigureAwait(false);
    }
}