using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using ReplyBrief.Desktop.ViewModels;

namespace ReplyBrief.Desktop;

public partial class MainWindow : Window
{
    // Below this the two columns no longer fit side by side, so the side column drops underneath.
    private const double StackingWidth = 860;

    private MainViewModel? _viewModel;
    private bool _allowClose;
    private bool _stacked;

    public MainWindow()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => _viewModel = DataContext as MainViewModel;
        SizeChanged += (_, args) => QueueLayout(args.NewSize.Width < StackingWidth);
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        base.OnClosing(e);

        var viewModel = _viewModel;
        if (_allowClose || viewModel is null)
            return;

        // Cancel the run in flight and close the capture host, then let the window go.
        e.Cancel = true;
        _allowClose = true;
        _ = ShutdownAsync(viewModel);
    }

    private async Task ShutdownAsync(MainViewModel viewModel)
    {
        // WPF refuses Close() while a window is mid-close, so leave the Closing call stack first.
        await Dispatcher.InvokeAsync(static () => { }, DispatcherPriority.ApplicationIdle);
        await viewModel.DisposeAsync();
        _ = Dispatcher.BeginInvoke(new Action(Close));
    }

    private void QueueLayout(bool stacked)
    {
        if (stacked == _stacked)
            return;

        _stacked = stacked;

        // Resizing columns from inside SizeChanged arranges children mid-measure and leaves them
        // misplaced, so the change waits for the layout pass that raised it to finish.
        _ = Dispatcher.BeginInvoke(new Action(() => ApplyLayout(stacked)), DispatcherPriority.Loaded);
    }

    private void ApplyLayout(bool stacked)
    {
        var conversation = LayoutGrid.ColumnDefinitions[0];
        var gutter = LayoutGrid.ColumnDefinitions[1];
        var side = LayoutGrid.ColumnDefinitions[2];

        conversation.MinWidth = stacked ? 0 : 380;
        conversation.Width = new GridLength(stacked ? 1 : 3, GridUnitType.Star);
        gutter.Width = new GridLength(stacked ? 0 : 16);
        side.MinWidth = stacked ? 0 : 260;
        side.Width = new GridLength(stacked ? 0 : 2, GridUnitType.Star);

        // Stacked, the side column moves under the conversation and the reading moves below both.
        Grid.SetColumn(SideColumn, stacked ? 0 : 2);
        Grid.SetRow(SideColumn, stacked ? 1 : 0);
        Grid.SetRow(ReadingCard, stacked ? 2 : 1);
    }
}
