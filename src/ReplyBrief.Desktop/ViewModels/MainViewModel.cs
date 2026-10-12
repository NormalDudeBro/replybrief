using System.ComponentModel;
using System.IO;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using ReplyBrief.Core;
using ReplyBrief.Desktop.Commands;
using ReplyBrief.Desktop.Services;

namespace ReplyBrief.Desktop.ViewModels;

internal enum CaptureOrigin
{
    Manual,
    LocalHost,
}

internal enum StatusKind
{
    Info,
    Busy,
    Success,
    Error,
}

internal sealed class MainViewModel : INotifyPropertyChanged
{
    private static readonly TimeSpan AnalysisTimeout = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan ConnectionTimeout = TimeSpan.FromSeconds(10);
    private const int MaxTagsBytes = 128 * 1024;
    private const string DefaultOllamaHost = "http://localhost:11434";
    private const string DefaultOllamaModel = "llama3.1";

    private static readonly string[] ResultProperties =
    [
        nameof(Summary),
        nameof(NeedsAttention),
        nameof(ReplyPoints),
        nameof(MissingInformation),
    ];

    private static readonly string[] DerivedProperties =
    [
        nameof(Title),
        nameof(Participants),
        nameof(ContextUsed),
        nameof(Content),
        nameof(ContentCountText),
        nameof(IsContentOverLimit),
        nameof(OriginLabel),
        nameof(CoverageNotice),
        nameof(IsCaptureEdited),
        nameof(IsCaptureEditable),
        nameof(IsSettingsEditable),
        nameof(NetworkConsentText),
        nameof(NetworkConsentEnabled),
        nameof(HasPendingCapture),
        nameof(PendingCaptureNotice),
        nameof(HasResults),
        nameof(HasSummary),
        nameof(HasNeedsAttention),
        nameof(HasReplyPoints),
        nameof(HasMissingInformation),
        nameof(HasAvailableModels),
        nameof(AvailableModelList),
    ];

    private readonly ConversationAnalyzer _analyzer;
    private readonly HttpClient _httpClient;
    private readonly Dispatcher _dispatcher;
    private readonly DispatcherTimer _pairingNoticeTimer;
    private readonly CancellationTokenSource _shutdown = new();

    private LocalCaptureService? _localCapture;
    private Task _captureStartup = Task.CompletedTask;
    private Task? _analysisTask;
    private Task? _connectionTask;
    private CancellationTokenSource? _analysisCancellation;
    private ConversationCapture? _pendingCapture;
    private bool _userCancelled;
    private bool _suppressEditTracking;
    private CaptureOrigin _origin = CaptureOrigin.Manual;
    private string _reportedSource = string.Empty;
    private string _source = "manual";

    private string _title = string.Empty;
    private string _participants = string.Empty;
    private string _contextUsed = string.Empty;
    private string _content = string.Empty;
    private string _ollamaHost;
    private string _model;
    private bool _networkConsent;
    private bool _isCheckingConnection;
    private string _availableModelList = string.Empty;
    private string _pairingCode = "Not available";
    private string _pairingNotice = string.Empty;
    private bool _isAnalyzing;
    private bool _isLocalCaptureRunning;
    private bool _localCaptureHintVisible;
    private string _statusText = "Preparing ReplyBrief.";
    private StatusKind _statusKind = StatusKind.Info;

    private string[] _summary = [];
    private string[] _needsAttention = [];
    private string[] _replyPoints = [];
    private string[] _missingInformation = [];

    public MainViewModel(ConversationAnalyzer analyzer, HttpClient httpClient, string? ollamaHost = null, string? ollamaModel = null)
    {
        _analyzer = analyzer;
        _httpClient = httpClient;
        _dispatcher = Dispatcher.CurrentDispatcher;
        _ollamaHost = TrimmedOr(ollamaHost, DefaultOllamaHost);
        _model = TrimmedOr(ollamaModel, DefaultOllamaModel);

        _pairingNoticeTimer = new DispatcherTimer(
            TimeSpan.FromSeconds(4),
            DispatcherPriority.Background,
            OnPairingNoticeElapsed,
            _dispatcher);

        AnalyzeCommand = new RelayCommand(
            () => _analysisTask = AnalyzeAsync(),
            () => !IsAnalyzing && !IsCheckingConnection && !string.IsNullOrWhiteSpace(Content));
        CancelCommand = new RelayCommand(CancelAnalysis, () => IsAnalyzing);
        ClearCommand = new RelayCommand(ClearCapture, () => !IsAnalyzing);
        CopyPairingCommand = new RelayCommand(CopyPairingCode, () => IsLocalCaptureRunning);
        CheckConnectionCommand = new RelayCommand(
            () => _connectionTask = CheckConnectionAsync(),
            () => !IsAnalyzing && !IsCheckingConnection);
        LoadPendingCommand = new RelayCommand(LoadPendingCapture, () => HasPendingCapture);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public RelayCommand AnalyzeCommand { get; }

    public RelayCommand CancelCommand { get; }

    public RelayCommand ClearCommand { get; }

    public RelayCommand CopyPairingCommand { get; }

    public RelayCommand CheckConnectionCommand { get; }

    public RelayCommand LoadPendingCommand { get; }

    public string Title
    {
        get => _title;
        set
        {
            if (!Set(ref _title, value))
                return;

            NoteEdit();
            ResetConsent();
            ClearResults();
        }
    }

    public string Participants
    {
        get => _participants;
        set
        {
            if (!Set(ref _participants, value))
                return;

            NoteEdit();
            ResetConsent();
            ClearResults();
        }
    }

    public string ContextUsed
    {
        get => _contextUsed;
        set
        {
            if (!Set(ref _contextUsed, value))
                return;

            NoteEdit();
            ResetConsent();
            ClearResults();
        }
    }

    public string Content
    {
        get => _content;
        set
        {
            if (!Set(ref _content, value))
                return;

            NoteEdit();
            ResetConsent();
            ClearResults();
        }
    }

    public string SourceText
    {
        get => _source;
        private set => Set(ref _source, value);
    }

    public string OllamaHost
    {
        get => _ollamaHost;
        set
        {
            if (!Set(ref _ollamaHost, value))
                return;

            ClearModelList();
            ResetConsent();
        }
    }

    public string Model
    {
        get => _model;
        set
        {
            if (!Set(ref _model, value))
                return;

            ResetConsent();
        }
    }

    public bool NetworkConsent
    {
        get => _networkConsent;
        set => Set(ref _networkConsent, value);
    }

    public bool NetworkConsentEnabled => !IsAnalyzing && !IsCheckingConnection;

    public string NetworkConsentText => $"Send the previewed conversation to this Ollama server: {OllamaHost}";

    public bool IsCheckingConnection
    {
        get => _isCheckingConnection;
        private set
        {
            if (Set(ref _isCheckingConnection, value))
                OnChanged(nameof(CheckConnectionLabel));
        }
    }

    public string CheckConnectionLabel => IsCheckingConnection ? "Checking..." : "Check connection";

    public string AvailableModelList => _availableModelList;

    public bool HasAvailableModels => !string.IsNullOrEmpty(_availableModelList);

    /// <summary>A capture arrived while a run was in flight and is waiting to be loaded on purpose.</summary>
    public bool HasPendingCapture => _pendingCapture is not null && !IsAnalyzing;

    public string PendingCaptureNotice =>
        _pendingCapture is null
            ? string.Empty
            : $"A capture from {_pendingCapture.Source} arrived while ReplyBrief was analyzing. Loading it replaces the text above.";

    public string PairingCode
    {
        get => _pairingCode;
        private set => Set(ref _pairingCode, value);
    }

    public string PairingNotice
    {
        get => _pairingNotice;
        private set
        {
            if (Set(ref _pairingNotice, value))
                OnChanged(nameof(HasPairingNotice));
        }
    }

    public bool HasPairingNotice => !string.IsNullOrEmpty(PairingNotice);

    public bool IsLocalCaptureRunning
    {
        get => _isLocalCaptureRunning;
        private set
        {
            if (Set(ref _isLocalCaptureRunning, value))
                OnChanged(nameof(LocalCaptureStatusText));
        }
    }

    public string LocalCaptureStatusText => IsLocalCaptureRunning
        ? $"Listening on {LocalCaptureService.Endpoint}"
        : "Not running. Manual capture below is unaffected.";

    public bool LocalCaptureHintVisible
    {
        get => _localCaptureHintVisible;
        private set => Set(ref _localCaptureHintVisible, value);
    }

    public string LocalCaptureHint { get; } =
        "The capture host listens over HTTPS on this machine, so it needs a trusted certificate. " +
        "Install the .NET SDK if needed, then run: dotnet dev-certs https --trust  and restart ReplyBrief. " +
        "Captured text stays in this window until you send it to the server you choose below.";

    public bool IsAnalyzing
    {
        get => _isAnalyzing;
        private set => Set(ref _isAnalyzing, value);
    }

    public bool IsCaptureEditable => !IsAnalyzing;

    public bool IsSettingsEditable => !IsAnalyzing && !IsCheckingConnection;

    public string StatusText
    {
        get => _statusText;
        private set => Set(ref _statusText, value);
    }

    public string StatusTone => _statusKind switch
    {
        StatusKind.Busy => "busy",
        StatusKind.Success => "success",
        StatusKind.Error => "error",
        _ => "info",
    };

    public string ContentCountText => $"{Content.Length:N0} of {ConversationCapture.MaxContentLength:N0} characters";

    public bool IsContentOverLimit => Content.Length > ConversationCapture.MaxContentLength;

    public bool IsCaptureEdited { get; private set; }

    public string OriginLabel => _origin == CaptureOrigin.LocalHost
        ? _reportedSource.Equals(_source, StringComparison.OrdinalIgnoreCase)
            ? "Captured by the local capture host"
            : $"Captured by the local capture host (reported source: {_reportedSource})"
        : "Entered manually in ReplyBrief";

    public string CoverageNotice => _origin switch
    {
        CaptureOrigin.LocalHost when IsCaptureEdited =>
            "Captured locally and then edited by you. ReplyBrief cannot tell whether earlier messages in the thread were included, and it never adds messages that are not in the text above.",
        CaptureOrigin.LocalHost =>
            "Captured by the local capture host. ReplyBrief cannot tell whether earlier messages in the thread were included, and it never adds messages that are not in the text above.",
        _ =>
            "You supplied this text. ReplyBrief has no knowledge of messages outside it and will not guess at earlier turns in the conversation.",
    };

    public bool HasResults => _summary.Length > 0 || _needsAttention.Length > 0 || _replyPoints.Length > 0 || _missingInformation.Length > 0;

    public bool HasSummary => _summary.Length > 0;

    public bool HasNeedsAttention => _needsAttention.Length > 0;

    public bool HasReplyPoints => _replyPoints.Length > 0;

    public bool HasMissingInformation => _missingInformation.Length > 0;

    public IReadOnlyList<string> Summary => _summary;

    public IReadOnlyList<string> NeedsAttention => _needsAttention;

    public IReadOnlyList<string> ReplyPoints => _replyPoints;

    public IReadOnlyList<string> MissingInformation => _missingInformation;

    /// <summary>Starts the capture host and keeps the task so shutdown can wait for it.</summary>
    public Task StartLocalCaptureAsync() => _captureStartup = StartLocalCaptureCoreAsync();

    public async ValueTask DisposeAsync()
    {
        _shutdown.Cancel();
        CancelAnalysis();
        _pairingNoticeTimer.Stop();

        // Every task below posts its continuations to this dispatcher, so awaiting them from the UI
        // thread yields rather than blocks. The capture host is closed last, once nothing can still
        // call back into a disposed window.
        await Settle(_captureStartup);
        await Settle(_analysisTask);
        await Settle(_connectionTask);
        await DisposeLocalCaptureAsync();
    }

    /// <summary>
    /// Takes a capture from the connector onto the UI thread. Internal so tests can drive the same
    /// path the background callback takes.
    /// </summary>
    internal void ReceiveCapture(ConversationCapture capture) =>
        _dispatcher.BeginInvoke(new Action(() => ApplyCapture(capture)));

    private async Task StartLocalCaptureCoreAsync()
    {
        SetStatus("Starting the local capture host.", StatusKind.Busy);

        LocalCaptureService? capture = null;
        try
        {
            capture = await LocalCaptureService.StartAsync(HandleCaptureFromHost, _shutdown.Token);

            // Shutdown can land between the host coming up and this line. Closing the host without
            // touching bound state is the only safe outcome: the window is already on its way out.
            if (_shutdown.IsCancellationRequested)
            {
                await capture.DisposeAsync();
                return;
            }

            _localCapture = capture;
            PairingCode = capture.PairingCode;
            IsLocalCaptureRunning = true;
            LocalCaptureHintVisible = false;
            SetStatus("Local capture is listening. Paste a conversation below, or send one to the capture host.", StatusKind.Info);
        }
        catch (OperationCanceledException)
        {
            await DisposeQuietly(capture);
        }
        catch (Exception)
        {
            // Deliberately broad: a missing dev certificate, a busy port or an unknown host failure
            // must not take manual capture down with it.
            await DisposeQuietly(capture);
            if (_shutdown.IsCancellationRequested)
                return;

            await DisposeLocalCaptureAsync();
            IsLocalCaptureRunning = false;
            PairingCode = "Not available";
            LocalCaptureHintVisible = true;
            SetStatus("Local capture could not start. Manual capture below still works.", StatusKind.Error);
        }
    }

    private async Task AnalyzeAsync()
    {
        if (IsAnalyzing || IsCheckingConnection)
            return;

        if (!NetworkConsent)
        {
            SetStatus("Confirm that the conversation may be sent to this Ollama server first.", StatusKind.Error);
            return;
        }

        var host = OllamaHost.Trim();
        ConversationCapture capture;
        try
        {
            capture = new ConversationCapture(
                _origin == CaptureOrigin.LocalHost ? _source : "manual",
                Title.Trim(),
                Participants.Trim(),
                Content,
                ContextUsed.Trim());
            capture.Validate();
        }
        catch (Exception ex) when (ex is ArgumentException or AnalysisException)
        {
            SetStatus(ex.Message, StatusKind.Error);
            return;
        }

        // A run that has not produced anything yet must never sit beside notes from an earlier one.
        ClearResults();

        using var timeout = new CancellationTokenSource(AnalysisTimeout);
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(timeout.Token, _shutdown.Token);
        _analysisCancellation = linked;
        _userCancelled = false;
        IsAnalyzing = true;
        SetStatus($"Sending the conversation to {host} and waiting for a reply.", StatusKind.Busy);

        try
        {
            ApplyResult(await _analyzer.AnalyzeAsync(capture, new AnalysisOptions(Model.Trim(), host), linked.Token));
            SetStatus(
                _pendingCapture is null
                    ? "Analysis complete."
                    : "Analysis complete. A capture arrived during the run and is waiting to be loaded.",
                StatusKind.Success);
        }
        catch (OperationCanceledException)
        {
            SetStatus(DescribeCancellation(host), StatusKind.Error);
        }
        catch (AnalysisException ex)
        {
            SetStatus(ex.Message, StatusKind.Error);
        }
        catch (HttpRequestException)
        {
            SetStatus($"Could not reach {host}. Check the address and that the server is running.", StatusKind.Error);
        }
        catch (Exception)
        {
            // Anything unexpected may carry model or API text, so it is reported without details.
            SetStatus("The analysis failed unexpectedly. Details are withheld because they can contain conversation text.", StatusKind.Error);
        }
        finally
        {
            _analysisCancellation = null;
            IsAnalyzing = false;
        }
    }

    private async Task CheckConnectionAsync()
    {
        if (IsAnalyzing || IsCheckingConnection)
            return;

        Uri endpoint;
        try
        {
            endpoint = ConversationAnalyzer.NormalizeEndpoint(OllamaHost);
        }
        catch (AnalysisException ex)
        {
            SetStatus(ex.Message, StatusKind.Error);
            return;
        }

        // The address and model are locked for the duration, and these snapshots prove it: a check
        // that answers late must not report against a server or model the window no longer shows.
        var requestedHost = OllamaHost;
        var requestedModel = Model.Trim();
        var tags = new Uri(endpoint, "api/tags");

        using var timeout = new CancellationTokenSource(ConnectionTimeout);
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(timeout.Token, _shutdown.Token);
        IsCheckingConnection = true;
        ClearModelList();
        SetStatus($"Asking {endpoint.Authority} which models it has.", StatusKind.Busy);

        try
        {
            using var response = await _httpClient.GetAsync(tags, HttpCompletionOption.ResponseHeadersRead, linked.Token);
            if (!response.IsSuccessStatusCode)
            {
                SetStatus($"The server at {endpoint.Authority} answered {(int)response.StatusCode}. Check the address and that it is an Ollama server.", StatusKind.Error);
                return;
            }

            var models = await ReadModelNamesAsync(response, linked.Token);

            if (OllamaHost != requestedHost || Model.Trim() != requestedModel)
            {
                SetStatus("The address or model changed while the check was running, so its answer was discarded.", StatusKind.Error);
                return;
            }

            PublishModels(models, requestedModel, endpoint.Authority);
        }
        catch (OperationCanceledException)
        {
            if (_shutdown.IsCancellationRequested)
                return;

            SetStatus($"{endpoint.Authority} did not answer within {ConnectionTimeout.TotalSeconds:0} seconds.", StatusKind.Error);
        }
        catch (AnalysisException ex)
        {
            SetStatus(ex.Message, StatusKind.Error);
        }
        catch (HttpRequestException)
        {
            SetStatus($"Could not reach {endpoint.Authority}. Check the address and that the server is running.", StatusKind.Error);
        }
        catch (Exception)
        {
            SetStatus("The connection check failed. Details are withheld because they can contain server text.", StatusKind.Error);
        }
        finally
        {
            IsCheckingConnection = false;
        }
    }

    private void PublishModels(IReadOnlyList<string> models, string requestedModel, string authority)
    {
        ClearModelList();
        _availableModelList = models.Count == 0
            ? "The server reported no models."
            : string.Join(", ", models);
        OnChanged(nameof(AvailableModelList));

        if (models.Count == 0)
        {
            SetStatus($"Connected to {authority}, but it reported no models.", StatusKind.Error);
            return;
        }

        var available = models.Contains(requestedModel, StringComparer.OrdinalIgnoreCase);
        SetStatus(
            available
                ? $"Connected to {authority}. {models.Count} models available, including {requestedModel}."
                : $"Connected to {authority}. {models.Count} models available, but {requestedModel} is not one of them.",
            available ? StatusKind.Success : StatusKind.Error);
    }

    private string DescribeCancellation(string host)
    {
        if (_shutdown.IsCancellationRequested)
            return "ReplyBrief is closing.";

        return _userCancelled
            ? "Analysis cancelled. The conversation stays in the window until you clear it."
            : $"{host} did not answer within {AnalysisTimeout.TotalMinutes:0} minutes. Check that it is reachable and that the model is pulled, then try a shorter conversation.";
    }

    private void CancelAnalysis()
    {
        if (!IsAnalyzing)
            return;

        _userCancelled = true;
        _analysisCancellation?.Cancel();
    }

    private void ClearCapture()
    {
        _suppressEditTracking = true;
        _title = string.Empty;
        _participants = string.Empty;
        _contextUsed = string.Empty;
        _content = string.Empty;
        _suppressEditTracking = false;

        _origin = CaptureOrigin.Manual;
        _reportedSource = string.Empty;
        SourceText = "manual";
        IsCaptureEdited = false;
        _pendingCapture = null;

        ResetConsent();
        ClearResults();
        RaiseDerived();
        SetStatus("Capture cleared. Nothing from the previous conversation is left in the window.", StatusKind.Info);
    }

    private void LoadPendingCapture()
    {
        if (!HasPendingCapture)
            return;

        var capture = _pendingCapture!;
        _pendingCapture = null;
        ApplyCapture(capture);
    }

    private void ApplyResult(AnalysisResult result)
    {
        _summary = Tidy(result.Summary);
        _needsAttention = Tidy(result.NeedsAttention);
        _replyPoints = Tidy(result.SuggestedReplyPoints);
        _missingInformation = Tidy(result.MissingInformation);

        NotifyResults();
        RaiseDerived();
    }

    private void ClearResults()
    {
        if (!HasResults)
            return;

        _summary = [];
        _needsAttention = [];
        _replyPoints = [];
        _missingInformation = [];
        NotifyResults();
        RaiseDerived();
    }

    private void HandleCaptureFromHost(ConversationCapture capture)
    {
        if (capture is null || _shutdown.IsCancellationRequested)
            return;

        try
        {
            ReceiveCapture(capture);
        }
        catch (InvalidOperationException)
        {
            // The dispatcher is shutting down; there is no window left to deliver the capture to.
        }
    }

    private void ApplyCapture(ConversationCapture capture)
    {
        if (IsAnalyzing)
        {
            // The host has already accepted this capture and cannot be told, so it is held rather
            // than dropped, and the newest one wins. It is never loaded over a result on its own.
            _pendingCapture = capture;
            RaiseDerived();
            SetStatus($"A capture from {capture.Source} arrived during the analysis. It is held until you load it, so it cannot be mixed with this result.", StatusKind.Error);
            return;
        }

        _reportedSource = capture.Source?.Trim() ?? string.Empty;
        _suppressEditTracking = true;
        _origin = CaptureOrigin.LocalHost;
        SourceText = NormalizeSource(_reportedSource);
        _title = capture.Title ?? string.Empty;
        _participants = capture.Participants ?? string.Empty;
        _contextUsed = capture.ContextUsed ?? string.Empty;
        _content = capture.Content ?? string.Empty;
        _suppressEditTracking = false;
        IsCaptureEdited = false;

        ResetConsent();
        ClearResults();
        RaiseDerived();
        SetStatus(
            IsContentOverLimit
                ? "Capture received. It is longer than the character limit, so trim it before analyzing."
                : "Capture received. Review it below, confirm the server, then run an analysis. Nothing is analyzed automatically.",
            StatusKind.Info);
    }

    private void CopyPairingCode()
    {
        if (!IsLocalCaptureRunning)
            return;

        try
        {
            Clipboard.SetText(PairingCode);
            PairingNotice = "Copied to the clipboard";
            _pairingNoticeTimer.Stop();
            _pairingNoticeTimer.Start();
        }
        catch (ExternalException)
        {
            SetStatus("The pairing code could not reach the clipboard. Select it and copy it by hand.", StatusKind.Error);
        }
    }

    private void OnPairingNoticeElapsed(object? sender, EventArgs e)
    {
        _pairingNoticeTimer.Stop();
        PairingNotice = string.Empty;
    }

    private void NoteEdit()
    {
        if (_suppressEditTracking || _origin != CaptureOrigin.LocalHost || IsCaptureEdited)
            return;

        IsCaptureEdited = true;
        RaiseDerived();
    }

    private void ResetConsent()
    {
        if (!NetworkConsent)
            return;

        NetworkConsent = false;
    }

    private void ClearModelList()
    {
        if (_availableModelList.Length == 0)
            return;

        _availableModelList = string.Empty;
        OnChanged(nameof(AvailableModelList));
    }

    private static async Task<IReadOnlyList<string>> ReadModelNamesAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        if (response.Content.Headers.ContentLength is > MaxTagsBytes)
            throw new AnalysisException("The server sent an oversized model list.");

        await using var body = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var buffer = new MemoryStream();
        var chunk = new byte[8192];
        int count;
        while ((count = await body.ReadAsync(chunk, cancellationToken)) > 0)
        {
            if (buffer.Length + count > MaxTagsBytes)
                throw new AnalysisException("The server sent an oversized model list.");

            buffer.Write(chunk, 0, count);
        }

        try
        {
            using var document = JsonDocument.Parse(buffer.ToArray());
            if (document.RootElement.ValueKind != JsonValueKind.Object ||
                !document.RootElement.TryGetProperty("models", out var models) ||
                models.ValueKind != JsonValueKind.Array)
                throw new AnalysisException("The server answered, but not with an Ollama model list.");

            return models.EnumerateArray()
                .Select(model => model.ValueKind == JsonValueKind.Object &&
                                 model.TryGetProperty("name", out var name) &&
                                 name.ValueKind == JsonValueKind.String
                    ? name.GetString()
                    : null)
                .OfType<string>()
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name.Trim())
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
        catch (JsonException)
        {
            throw new AnalysisException("The server's model list was not readable JSON.");
        }
    }

    private static string TrimmedOr(string? value, string fallback) =>
        string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();

    private static string NormalizeSource(string reported)
    {
        var source = reported.ToLowerInvariant();
        return source is "outlook" or "gmail" or "teams" or "manual" ? source : "manual";
    }

    private static string[] Tidy(string[]? items) =>
        (items ?? []).Where(item => !string.IsNullOrWhiteSpace(item)).Select(item => item.Trim()).ToArray();

    private void SetStatus(string message, StatusKind kind)
    {
        _statusKind = kind;
        StatusText = message;
        OnChanged(nameof(StatusTone));
    }

    private async Task DisposeLocalCaptureAsync()
    {
        var capture = _localCapture;
        if (capture is null)
            return;

        _localCapture = null;
        await DisposeQuietly(capture);
    }

    private static async Task DisposeQuietly(LocalCaptureService? capture)
    {
        if (capture is null)
            return;

        try
        {
            await capture.DisposeAsync();
        }
        catch (Exception)
        {
            // The window is closing; there is nowhere left to report a failed shutdown.
        }
    }

    private static async Task Settle(Task? task)
    {
        if (task is null)
            return;

        try
        {
            await task;
        }
        catch (Exception)
        {
            // The operation reports its own failure in the status line.
        }
    }

    private bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;

        field = value;
        OnChanged(name);
        RaiseDerived();
        CommandManager.InvalidateRequerySuggested();
        return true;
    }

    private void NotifyResults()
    {
        foreach (var name in ResultProperties)
            OnChanged(name);
    }

    private void RaiseDerived()
    {
        foreach (var name in DerivedProperties)
            OnChanged(name);
    }

    private void OnChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
