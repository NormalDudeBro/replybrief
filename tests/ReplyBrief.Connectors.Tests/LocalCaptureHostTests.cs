using System.Net;
using System.Net.NetworkInformation;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;

namespace ReplyBrief.Connectors.Tests;

public sealed class LocalCaptureHostTests
{
    [Fact]
    public async Task AcceptsAWellFormedOutlookCaptureAndHandsItToTheDesktopCallback()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), server.PairingCode);

        Assert.Equal(HttpStatusCode.Accepted, response.StatusCode);
        Assert.Equal("""{"status":"accepted"}""", await response.Content.ReadAsStringAsync());

        var capture = await recorder.NextAsync();

        Assert.Equal("outlook", capture.Source);
        Assert.Equal("Project sync", capture.Title);
        Assert.Equal("Dana Example <dana@example.com>", capture.Participants);
        Assert.Equal("Hello, are we still on for Thursday?", capture.Content);
        Assert.Equal("Outlook reading pane", capture.ContextUsed);
    }

    [Fact]
    public async Task AnswersTheTaskPaneBeforeTheDesktopCallbackFinishes()
    {
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        await using var server = await LoopbackCaptureServer.StartAsync(_ =>
        {
            entered.TrySetResult();
            release.Task.GetAwaiter().GetResult();
        });

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), server.PairingCode);

        Assert.Equal(HttpStatusCode.Accepted, response.StatusCode);

        await entered.Task.WaitAsync(TimeSpan.FromSeconds(15));
        release.SetResult();
    }

    [Fact]
    public async Task AcceptsARequestThatCarriesItsOwnOrigin()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(
            CaptureRecorder.OutlookCaptureJson(),
            server.PairingCode,
            origin: server.BaseAddress.GetLeftPart(UriPartial.Authority));

        Assert.Equal(HttpStatusCode.Accepted, response.StatusCode);
        Assert.NotNull(await recorder.NextAsync());
    }

    [Fact]
    public async Task RejectsAForeignOriginEvenWithAValidPairingCode()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), server.PairingCode, origin: "https://attacker.example");

        Assert.Equal(HttpStatusCode.Forbidden, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Fact]
    public async Task RejectsAPlainHttpOriginEvenOnTheRightAuthority()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), server.PairingCode, origin: $"http://localhost:{server.Port}");

        Assert.Equal(HttpStatusCode.Forbidden, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("obviously-not-the-pairing-code")]
    [InlineData("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")]
    public async Task RejectsAMissingOrWrongPairingCode(string? presented)
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), presented);

        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Fact]
    public async Task RejectsAPairingCodeFromAPreviousRun()
    {
        var previous = await LoopbackCaptureServer.StartAsync(_ => { });
        var staleCode = previous.PairingCode;
        await previous.DisposeAsync();

        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), staleCode);

        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Theory]
    [InlineData("text/plain")]
    [InlineData("application/x-www-form-urlencoded")]
    [InlineData("text/json")]
    [InlineData("application/xml")]
    [InlineData(null)]
    public async Task RejectsACaptureThatIsNotSentAsJson(string? contentType)
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), server.PairingCode, contentType);

        Assert.Equal(HttpStatusCode.UnsupportedMediaType, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Fact]
    public async Task RejectsACaptureLargerThanTheRequestLimit()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        var oversized = CaptureRecorder.OutlookCaptureJson(new string('x', LocalCaptureHost.MaxRequestBytes));
        Assert.True(Encoding.UTF8.GetByteCount(oversized) > LocalCaptureHost.MaxRequestBytes);

        using var response = await server.PostCaptureAsync(oversized, server.PairingCode);

        Assert.Equal(HttpStatusCode.RequestEntityTooLarge, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Theory]
    [InlineData("not json at all")]
    [InlineData("{ \"source\": \"outlook\", ")]
    [InlineData("[]")]
    [InlineData("\"outlook\"")]
    [InlineData("{ \"source\": \"outlook\", \"title\": \"a\", \"content\": \"x\" } trailing")]
    public async Task RejectsABodyThatIsNotAWellFormedCaptureObject(string body)
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(body, server.PairingCode);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Theory]
    [InlineData("gmail")]
    [InlineData("manual")]
    [InlineData("Outlook")]
    [InlineData("")]
    public async Task RejectsASourceOtherThanOutlook(string source)
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(source: source), server.PairingCode);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public async Task RejectsAnEmptyConversation(string content)
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(content), server.PairingCode);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Fact]
    public async Task RejectsAConversationThatExceedsTheCoreContentLimit()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(new string('x', 60_001)), server.PairingCode);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Fact]
    public async Task RejectsAFieldOfTheWrongTypeInsteadOfCoercingIt()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.PostCaptureAsync(
            """{ "source": "outlook", "content": 42, "title": "Project sync" }""",
            server.PairingCode);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Fact]
    public async Task RejectsGetOnTheCaptureEndpoint()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.GetAsync(LocalCaptureHost.CapturePath);

        Assert.Equal(HttpStatusCode.MethodNotAllowed, response.StatusCode);
    }

    [Fact]
    public async Task ServesTheTaskPaneWithAStrictPolicyAndWithoutAnySecret()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.GetAsync("/outlook/taskpane.html");
        var html = await response.Content.ReadAsStringAsync();

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal("text/html; charset=utf-8", response.Content.Headers.ContentType?.ToString());
        Assert.Contains("https://appsforoffice.microsoft.com/lib/1/hosted/office.js", html);
        Assert.Contains("taskpane.js", html);

        var policy = Assert.Single(response.Headers.GetValues("Content-Security-Policy"));
        Assert.Contains("default-src 'none'", policy);
        Assert.Contains("script-src 'self' https://appsforoffice.microsoft.com", policy);
        Assert.Contains("frame-ancestors 'self'", policy);
        Assert.Contains("https://outlook.office.com", policy);
        Assert.Contains("https://*.outlook.com", policy);

        Assert.Equal("nosniff", Assert.Single(response.Headers.GetValues("X-Content-Type-Options")));
        Assert.DoesNotContain(server.PairingCode, html);
        Assert.DoesNotContain(LocalCaptureHost.PairingHeaderName, html);
    }

    [Fact]
    public async Task ServesTheManifestTheScriptAndTheIcons()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var script = await server.GetAsync("/outlook/taskpane.js");
        var scriptBody = await script.Content.ReadAsStringAsync();

        using var manifest = await server.GetAsync("/outlook/manifest.xml");

        using var icon = await server.GetAsync("/outlook/icon.svg");

        using var commands = await server.GetAsync("/outlook/commands.html");
        var commandsBody = await commands.Content.ReadAsStringAsync();

        using var support = await server.GetAsync("/outlook/support.html");

        Assert.Equal(HttpStatusCode.OK, script.StatusCode);
        Assert.Equal("text/javascript; charset=utf-8", script.Content.Headers.ContentType?.ToString());
        Assert.Contains($"\"{LocalCaptureHost.PairingHeaderName}\"", scriptBody);
        Assert.Contains("Office.CoercionType.Text", scriptBody);
        Assert.DoesNotContain(server.PairingCode, scriptBody);

        Assert.Equal(HttpStatusCode.OK, manifest.StatusCode);
        Assert.Equal("application/xml; charset=utf-8", manifest.Content.Headers.ContentType?.ToString());
        Assert.Equal(
            await File.ReadAllTextAsync(Path.Combine(AddInAssets.DefaultDirectory, "manifest.xml")),
            await manifest.Content.ReadAsStringAsync());

        Assert.Equal(HttpStatusCode.OK, icon.StatusCode);
        Assert.Equal("image/svg+xml", icon.Content.Headers.ContentType?.ToString());

        // The function file is a real page of its own that loads office.js and nothing else, so the
        // capture script can never be pulled into a context that has no UI and no pairing code.
        Assert.Equal(HttpStatusCode.OK, commands.StatusCode);
        Assert.Contains("office.js", commandsBody);
        Assert.DoesNotContain("src=\"taskpane.js\"", commandsBody);
        Assert.DoesNotContain("X-ReplyBrief-Pairing", commandsBody);

        Assert.Equal(HttpStatusCode.OK, support.StatusCode);
        Assert.Equal("text/html; charset=utf-8", support.Content.Headers.ContentType?.ToString());
    }

    [Theory]
    [InlineData("/outlook/manifest.json")]
    [InlineData("/outlook/app.js")]
    [InlineData("/outlook/taskpane.html.bak")]
    [InlineData("/outlook/commands.html.bak")]
    [InlineData("/outlook/..%2fappsettings.json")]
    [InlineData("/outlook/%2e%2e/%2e%2e/appsettings.json")]
    [InlineData("/outlook/subdir/taskpane.html")]
    [InlineData("/")]
    [InlineData("/capture")]
    public async Task ServesNothingOutsideTheAllowList(string path)
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var response = await server.GetAsync(path);

        Assert.True(
            response.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.BadRequest or HttpStatusCode.MethodNotAllowed,
            $"GET {path} returned {(int)response.StatusCode}");

        Assert.DoesNotContain(server.PairingCode, await response.Content.ReadAsStringAsync());
    }

    [Fact]
    public async Task ListensOnLoopbackOnly()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        var external = NetworkInterface.GetAllNetworkInterfaces()
            .Where(adapter => adapter.OperationalStatus == OperationalStatus.Up && adapter.NetworkInterfaceType != NetworkInterfaceType.Loopback)
            .SelectMany(adapter => adapter.GetIPProperties().UnicastAddresses)
            .Select(address => address.Address)
            .FirstOrDefault(address => address.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(address));

        if (external is null) return;

        using var client = new TcpClient();
        var connected = false;

        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            await client.ConnectAsync(external, server.Port, timeout.Token);
            connected = true;
        }
        catch (Exception) when (!connected)
        {
        }

        Assert.False(connected, $"the capture host accepted a connection on {external}, which is not a loopback address");
    }

    [Fact]
    public async Task HasNoPlainHttpFallbackOnTheCapturePort()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        HttpResponseMessage? response = null;

        try
        {
            response = await client.PostAsync(
                new Uri($"http://localhost:{server.Port}{LocalCaptureHost.CapturePath}"),
                new StringContent(CaptureRecorder.OutlookCaptureJson(), Encoding.UTF8, "application/json"));
        }
        catch (HttpRequestException)
        {
            // Kestrel answers a plaintext request aimed at a TLS endpoint by dropping the
            // connection. There is no HTTP listener for it to be routed to, which is the point.
        }

        try
        {
            Assert.False(response?.IsSuccessStatusCode ?? false);
        }
        finally
        {
            response?.Dispose();
        }

        Assert.Equal(0, recorder.Count);
    }

    [Fact]
    public async Task KeepsAcceptingCapturesAfterTheDesktopCallbackThrows()
    {
        var recorder = new CaptureRecorder();

        await using var server = await LoopbackCaptureServer.StartAsync(capture =>
        {
            if (capture.Title == "explode") throw new InvalidOperationException("desktop handler failed");
            recorder.Record(capture);
        });

        using var failing = await server.PostCaptureAsync(
            CaptureRecorder.OutlookCaptureJson().Replace("Project sync", "explode", StringComparison.Ordinal),
            server.PairingCode);

        using var following = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), server.PairingCode);

        Assert.Equal(HttpStatusCode.Accepted, failing.StatusCode);
        Assert.Equal(HttpStatusCode.Accepted, following.StatusCode);

        await recorder.NextAsync();
        Assert.Contains(server.DispatchFailures, failure => failure.Message == "desktop handler failed");
    }

    [Fact]
    public async Task ListensOnBothLoopbackFamiliesWhenBoundToAFixedPort()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record, LoopbackCaptureServer.FreeLoopbackPort());

        foreach (var family in new[] { "127.0.0.1", "[::1]" })
        {
            var origin = new Uri($"https://localhost:{server.Port}");
            using var response = await LoopbackCaptureServer.SendToAsync(
                origin,
                new HttpRequestMessage(HttpMethod.Get, new Uri($"https://{family}:{server.Port}/outlook/manifest.xml")));

            Assert.True(response.IsSuccessStatusCode, $"{family} was not reachable: {(int)response.StatusCode}");
        }
    }

    [Fact]
    public async Task RefusesCapturesOnceTheQueueIsFullInsteadOfGrowingWithoutBound()
    {
        // The desktop callback blocks, so nothing drains the queue while these captures arrive.
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        await using var server = await LoopbackCaptureServer.StartAsync(_ =>
        {
            entered.TrySetResult();
            release.Task.GetAwaiter().GetResult();
        });

        try
        {
            // One capture occupies the single dispatch slot and the rest fill the bounded queue.
            var statuses = new List<HttpStatusCode>();
            for (var attempt = 0; attempt < LocalCaptureHost.CaptureQueueCapacity + 4; attempt++)
                statuses.Add((await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), server.PairingCode)).StatusCode);

            await entered.Task.WaitAsync(TimeSpan.FromSeconds(15));

            // One capture sits with the blocked desktop callback and the rest fill the bounded
            // queue, so the host holds a fixed number and refuses the rest rather than growing.
            var accepted = statuses.Count(status => status == HttpStatusCode.Accepted);

            Assert.Equal(LocalCaptureHost.CaptureQueueCapacity + 1, accepted);
            Assert.Equal(HttpStatusCode.TooManyRequests, statuses[^1]);
            Assert.Equal(statuses.Count - accepted, statuses.Count(status => status == HttpStatusCode.TooManyRequests));
        }
        finally
        {
            release.SetResult();
        }
    }

    [Fact]
    public async Task RefusesAChunkedCaptureThatGrowsPastTheRequestLimit()
    {
        var recorder = new CaptureRecorder();
        await using var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        // PushStreamContent cannot compute a length, so the request is sent chunked and the limit can
        // only be enforced while the body streams in.
        var chunk = System.Text.Encoding.UTF8.GetBytes(new string('x', 16 * 1024));
        var request = new HttpRequestMessage(HttpMethod.Post, new Uri(server.BaseAddress, LocalCaptureHost.CapturePath))
        {
            Content = new StreamedCaptureContent(async stream =>
            {
                await stream.WriteAsync(Encoding.UTF8.GetBytes("{\"source\":\"outlook\",\"content\":\""), CancellationToken.None);
                await stream.FlushAsync(CancellationToken.None);

                for (var written = 0; written <= LocalCaptureHost.MaxRequestBytes; written += chunk.Length)
                {
                    await stream.WriteAsync(chunk, CancellationToken.None);
                    await stream.FlushAsync(CancellationToken.None);
                }

                await stream.WriteAsync(Encoding.UTF8.GetBytes("\"}"), CancellationToken.None);
            }),
        };
        request.Headers.TryAddWithoutValidation(LocalCaptureHost.PairingHeaderName, server.PairingCode);
        request.Content.Headers.TryAddWithoutValidation("Content-Type", "application/json");

        using var response = await LoopbackCaptureServer.SendToAsync(server.BaseAddress, request);

        Assert.Equal(HttpStatusCode.RequestEntityTooLarge, response.StatusCode);
        Assert.Equal(0, recorder.Count);
    }

    [Fact]
    public async Task StopsAcceptingCapturesOnceItHasBegunShuttingDown()
    {
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        var server = await LoopbackCaptureServer.StartAsync(_ =>
        {
            entered.TrySetResult();
            release.Task.GetAwaiter().GetResult();
        });

        using (var first = await server.PostCaptureAsync(CaptureRecorder.OutlookCaptureJson(), server.PairingCode))
        {
            Assert.Equal(HttpStatusCode.Accepted, first.StatusCode);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(15));

            // Start disposal while a capture is still blocked in the desktop callback, then release it
            // so the listener gets a chance to answer one more request before it goes away.
            var disposal = server.DisposeAsync().AsTask();
            await Task.Delay(50);
            release.SetResult();

            try
            {
                using var duringShutdown = await LoopbackCaptureServer.SendToAsync(server.BaseAddress, server.CaptureRequest(CaptureRecorder.OutlookCaptureJson()));

                Assert.True(
                    duringShutdown.StatusCode is HttpStatusCode.ServiceUnavailable or HttpStatusCode.Gone,
                    $"a capture during shutdown returned {(int)duringShutdown.StatusCode}");
            }
            catch (HttpRequestException)
            {
                // Equally acceptable: the listener may already have stopped accepting connections.
            }

            await disposal;
        }
    }

    [Fact]
    public async Task RotatesThePairingCodeAndThePortOnEveryStart()
    {
        var recorder = new CaptureRecorder();

        var first = await LoopbackCaptureServer.StartAsync(recorder.Record);
        var firstCode = first.PairingCode;
        var firstPort = first.Port;
        await first.DisposeAsync();

        await using var second = await LoopbackCaptureServer.StartAsync(recorder.Record);

        Assert.NotEqual(firstCode, second.PairingCode);
        Assert.NotEqual(firstPort, second.Port);
        Assert.Equal($"https://localhost:{second.Port}/", second.BaseAddress.AbsoluteUri);
    }

    [Fact]
    public async Task ToleratesBeingDisposedTwice()
    {
        var recorder = new CaptureRecorder();
        var server = await LoopbackCaptureServer.StartAsync(recorder.Record);

        await server.DisposeAsync();
        await server.DisposeAsync();
    }

    /// <summary>
    /// An <see cref="HttpContent"/> that reports no length, so HttpClient sends it with chunked
    /// transfer encoding. That is the only way to exercise a request whose size the server only
    /// discovers while reading, rather than from a Content-Length header.
    /// </summary>
    private sealed class StreamedCaptureContent(Func<Stream, Task> write) : HttpContent
    {
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) => write(stream);

        protected override bool TryComputeLength(out long length)
        {
            length = 0;
            return false;
        }
    }
}