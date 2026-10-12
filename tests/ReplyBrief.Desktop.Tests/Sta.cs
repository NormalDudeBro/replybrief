using System.Windows.Threading;

namespace ReplyBrief.Desktop.Tests;

/// <summary>
/// Runs a test body on its own STA thread with a live <see cref="Dispatcher"/>, which is the seam
/// <c>MainViewModel</c> is built against: it captures <c>Dispatcher.CurrentDispatcher</c> and every
/// command resumes through it. Awaiting here yields back to that dispatcher, exactly as the window
/// does, so no test needs a window and nothing deadlocks on a blocking wait.
/// </summary>
internal static class Sta
{
    public static Task RunAsync(Func<Task> body)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            var dispatcher = Dispatcher.CurrentDispatcher;
            _ = RunBodyAsync(dispatcher, body, completion);
            Dispatcher.Run();
        })
        {
            IsBackground = true,
            Name = "ReplyBrief desktop test",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private static async Task RunBodyAsync(Dispatcher dispatcher, Func<Task> body, TaskCompletionSource completion)
    {
        try
        {
            await body();
            completion.TrySetResult();
        }
        catch (Exception ex)
        {
            completion.TrySetException(ex);
        }
        finally
        {
            await dispatcher.InvokeAsync(static () => { }, DispatcherPriority.ApplicationIdle);
            dispatcher.InvokeShutdown();
        }
    }
}
