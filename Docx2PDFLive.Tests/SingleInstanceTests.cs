using System.Threading;
using Xunit;

namespace Docx2PDFLive.Tests;

public class SingleInstanceTests
{
    [Fact]
    public void NamedMutex_BlocksDuplicateInstance()
    {
        var testMutexName = @"Local\Test_Docx2PDF_SingleInstance_" + Guid.NewGuid().ToString("N");

        using var firstMutex = new Mutex(true, testMutexName, out bool firstAcquired);
        Assert.True(firstAcquired, "First instance must acquire the mutex.");

        using var secondMutex = new Mutex(true, testMutexName, out bool secondAcquired);
        Assert.False(secondAcquired, "Second clone instance must NOT acquire the mutex.");
    }

    [Fact]
    public async Task EventWaitHandle_SignalsWakeupToPrimaryInstance()
    {
        var testEventName = @"Local\Test_Docx2PDF_WakeEvent_" + Guid.NewGuid().ToString("N");

        using var primaryWaitHandle = new EventWaitHandle(false, EventResetMode.AutoReset, testEventName);

        var tcs = new TaskCompletionSource<bool>();

        var registeredWait = ThreadPool.RegisterWaitForSingleObject(
            primaryWaitHandle,
            (state, timedOut) =>
            {
                if (!timedOut)
                {
                    tcs.TrySetResult(true);
                }
            },
            null,
            TimeSpan.FromSeconds(5),
            false);

        try
        {
            // Simulate second clone instance waking up the primary instance
            using (var secondarySignal = EventWaitHandle.OpenExisting(testEventName))
            {
                secondarySignal.Set();
            }

            var completed = await Task.WhenAny(tcs.Task, Task.Delay(3000));
            Assert.True(completed == tcs.Task, "Primary instance must receive wake signal from second instance.");
        }
        finally
        {
            registeredWait.Unregister(null);
        }
    }
}
