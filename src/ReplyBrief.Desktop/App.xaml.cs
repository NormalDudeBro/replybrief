using System.Net.Http;
using System.Windows;
using ReplyBrief.Core;
using ReplyBrief.Desktop.ViewModels;

namespace ReplyBrief.Desktop;

public partial class App : Application
{
    // One client for the process. Redirects are refused so an Ollama server cannot bounce a
    // conversation to somewhere else, and the analysis timeout belongs to the view model.
    private readonly HttpClient _httpClient = new(new HttpClientHandler { AllowAutoRedirect = false })
    {
        Timeout = Timeout.InfiniteTimeSpan,
    };

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var viewModel = new MainViewModel(
            new ConversationAnalyzer(_httpClient),
            _httpClient,
            Environment.GetEnvironmentVariable("REPLYBRIEF_OLLAMA_URL"),
            Environment.GetEnvironmentVariable("REPLYBRIEF_OLLAMA_MODEL"));

        var window = new MainWindow { DataContext = viewModel };
        MainWindow = window;
        window.Show();
        _ = StartLocalCaptureAsync(viewModel);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _httpClient.Dispose();
        base.OnExit(e);
    }

    private static async Task StartLocalCaptureAsync(MainViewModel viewModel)
    {
        try
        {
            await viewModel.StartLocalCaptureAsync();
        }
        catch (Exception)
        {
            // The view model already reports capture startup problems in its status line.
        }
    }
}
