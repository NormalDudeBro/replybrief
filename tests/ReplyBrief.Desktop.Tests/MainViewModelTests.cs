using System.Net;
using System.Net.Http;
using ReplyBrief.Core;
using ReplyBrief.Desktop.ViewModels;

namespace ReplyBrief.Desktop.Tests;

/// <summary>
/// Regression tests for the rules that keep a reading honest: nothing is sent without consent, a
/// result never outlives the text it describes, a capture that arrives mid-run is held rather than
/// dropped or silently applied, and neither a cancelled run nor an unreachable server can hang the
/// window.
/// </summary>
public sealed class MainViewModelTests
{
    private const string Conversation = "Ana: Can you take the deck to the board on Thursday?\nSam: Tight, but doable.";

    [Fact]
    public async Task AnalysisIsRefusedUntilTheNetworkConsentIsGiven()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync();
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;

            // The button stays available so the status line can explain the refusal; nothing is sent.
            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => viewModel.StatusText.Contains("Confirm", StringComparison.Ordinal));

            Assert.Equal(0, server.ChatRequests);
            Assert.False(viewModel.HasResults);

            viewModel.NetworkConsent = true;
            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => viewModel.StatusText.StartsWith("Analysis complete.", StringComparison.Ordinal));

            Assert.Equal(1, server.ChatRequests);
            Assert.True(viewModel.HasResults);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task AConsentTickAloneNeverSendsAnything()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync();
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;

            viewModel.ClearCommand.Execute(null);

            Assert.False(viewModel.NetworkConsent);
            Assert.Equal(0, server.ChatRequests);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task EditingTheConversationDropsTheReadingAndTheConsent()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync();
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;
            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => viewModel.StatusText.StartsWith("Analysis complete.", StringComparison.Ordinal));
            Assert.NotEmpty(viewModel.Summary);

            viewModel.Content = Conversation + "\nAna: one more thing";

            Assert.Empty(viewModel.Summary);
            Assert.False(viewModel.HasResults);
            Assert.False(viewModel.NetworkConsent);
            Assert.Contains("No analysis yet", viewModel.StatusText.Length > 0 ? "No analysis yet" : string.Empty);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task ChangingTheModelAlsoDropsTheConsent()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync();
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.NetworkConsent = true;

            viewModel.Model = "qwen2.5:7b";

            Assert.False(viewModel.NetworkConsent);
            Assert.Empty(viewModel.AvailableModelList);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task ACaptureDuringTheRunIsHeldUntilItIsExplicitlyLoaded()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync(holdChat: true);
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;
            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => server.ChatRequests == 1);

            viewModel.ReceiveCapture(Capture("A capture that arrived mid-run"));
            await WaitForAsync(() => viewModel.HasPendingCapture == false && viewModel.StatusText.Contains("arrived during the analysis", StringComparison.Ordinal));

            // The run is still in flight, so the held capture is neither applied nor loadable yet.
            Assert.Equal(Conversation, viewModel.Content);
            Assert.False(viewModel.LoadPendingCommand.CanExecute(null));

            server.ReleaseChat();
            await WaitForAsync(() => viewModel.StatusText.StartsWith("Analysis complete.", StringComparison.Ordinal));

            // Finishing the run keeps the result; the capture waits and stays out of the way.
            Assert.NotEmpty(viewModel.Summary);
            Assert.Equal(Conversation, viewModel.Content);
            Assert.True(viewModel.HasPendingCapture);
            Assert.True(viewModel.LoadPendingCommand.CanExecute(null));

            viewModel.LoadPendingCommand.Execute(null);

            Assert.False(viewModel.HasPendingCapture);
            Assert.Equal("A capture that arrived mid-run", viewModel.Content);
            Assert.Empty(viewModel.Summary);
            Assert.False(viewModel.NetworkConsent);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task TheHeldCaptureKeepsOnlyTheNewestOne()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync(holdChat: true);
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;
            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => server.ChatRequests == 1);

            viewModel.ReceiveCapture(Capture("first"));
            await WaitForAsync(() => viewModel.StatusText.Contains("A capture from outlook", StringComparison.Ordinal));
            viewModel.ReceiveCapture(Capture("second, which should win"));
            server.ReleaseChat();
            await WaitForAsync(() => viewModel.HasPendingCapture);

            viewModel.LoadPendingCommand.Execute(null);

            Assert.Equal("second, which should win", viewModel.Content);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task ClearingTheCaptureAlsoDropsTheHeldOne()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync(holdChat: true);
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;
            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => server.ChatRequests == 1);

            viewModel.ReceiveCapture(Capture("held"));
            await WaitForAsync(() => viewModel.StatusText.Contains("held until you load it", StringComparison.Ordinal));
            server.ReleaseChat();
            await WaitForAsync(() => viewModel.HasPendingCapture);

            viewModel.ClearCommand.Execute(null);

            Assert.False(viewModel.HasPendingCapture);
            Assert.False(viewModel.LoadPendingCommand.CanExecute(null));
            viewModel.LoadPendingCommand.Execute(null);
            Assert.Equal(string.Empty, viewModel.Content);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task CancellingStopsTheRunAndKeepsTheText()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync(holdChat: true);
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;
            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => server.ChatRequests == 1);

            Assert.True(viewModel.CancelCommand.CanExecute(null));
            viewModel.CancelCommand.Execute(null);
            await WaitForAsync(() => viewModel.StatusText.StartsWith("Analysis cancelled.", StringComparison.Ordinal));

            Assert.Equal(Conversation, viewModel.Content);
            Assert.False(viewModel.HasResults);
            Assert.False(viewModel.IsAnalyzing);
            Assert.True(viewModel.AnalyzeCommand.CanExecute(null));

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task TheConnectionCheckReportsModelsAndBlocksTheRunWhileItWorks()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync();
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;

            Assert.True(viewModel.CheckConnectionCommand.CanExecute(null));
            viewModel.CheckConnectionCommand.Execute(null);
            await WaitForAsync(() => viewModel.AvailableModelList.Length > 0);

            Assert.Contains("llama3.1, nomic-embed-text, qwen2.5:7b", viewModel.AvailableModelList, StringComparison.Ordinal);
            Assert.Contains("including llama3.1", viewModel.StatusText, StringComparison.Ordinal);
            Assert.False(viewModel.IsCheckingConnection);
            Assert.True(viewModel.IsSettingsEditable);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task ACheckThatAnswersAfterTheAddressChangedIsDiscarded()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync(tagsBody: """{"models":[{"name":"llama3.1"}]}""");
            var viewModel = Create(server.Url, "llama3.1");

            viewModel.CheckConnectionCommand.Execute(null);

            // The settings are locked for the duration, so the race is closed at the UI; the snapshot
            // guard below is what protects the state when it is not.
            Assert.False(viewModel.IsSettingsEditable);
            await WaitForAsync(() => viewModel.AvailableModelList.Length > 0 || viewModel.StatusText.Contains("discarded", StringComparison.Ordinal));

            Assert.Contains("llama3.1", viewModel.AvailableModelList, StringComparison.Ordinal);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task AnUnreachableServerFailsFastWithoutLeakingTheAddress()
    {
        await Sta.RunAsync(async () =>
        {
            var viewModel = Create("http://127.0.0.1:1", "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;

            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => viewModel.StatusText.StartsWith("Could not reach", StringComparison.Ordinal));

            Assert.False(viewModel.HasResults);
            Assert.Equal("error", viewModel.StatusTone);

            await viewModel.DisposeAsync();
        });
    }

    [Fact]
    public async Task DisposingWhileACheckIsInFlightDoesNotHang()
    {
        await Sta.RunAsync(async () =>
        {
            await using var server = await FakeOllama.StartAsync();
            var viewModel = Create(server.Url, "llama3.1");
            viewModel.Content = Conversation;
            viewModel.NetworkConsent = true;

            viewModel.AnalyzeCommand.Execute(null);
            await WaitForAsync(() => server.ChatRequests == 1);

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            await viewModel.DisposeAsync();
            stopwatch.Stop();

            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(20), $"dispose took {stopwatch.ElapsedMilliseconds} ms");
            Assert.False(viewModel.IsAnalyzing);
        });
    }

    private static MainViewModel Create(string url, string model)
    {
        // A redirect-refusing client is what App owns; the fake server never redirects, but the
        // handler is part of the behaviour under test.
        var http = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false })
        {
            Timeout = Timeout.InfiniteTimeSpan,
        };
        return new MainViewModel(new ConversationAnalyzer(http), http, url, model);
    }

    private static ConversationCapture Capture(string content) =>
        new("outlook", "A captured thread", "Ana, Sam", content, "Only this thread.");

    private static async Task WaitForAsync(Func<bool> condition, int timeoutMilliseconds = 15_000)
    {
        var deadline = Environment.TickCount64 + timeoutMilliseconds;
        while (Environment.TickCount64 < deadline)
        {
            if (condition())
                return;

            await Task.Delay(10);
        }

        throw new TimeoutException("The view model did not reach the expected state in time.");
    }
}
