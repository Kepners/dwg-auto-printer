using System;
using System.Runtime.InteropServices;

namespace DwgAutoPrinter.App.Services;

internal static class ComMessageFilter
{
    public static IDisposable Register()
    {
        var filter = new MessageFilter();
        var hr = CoRegisterMessageFilter(filter, out var previous);
        if (hr < 0)
        {
            return NoopDisposable.Instance;
        }

        return new Registration(previous);
    }

    [DllImport("Ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    [ComImport]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    private sealed class MessageFilter : IOleMessageFilter
    {
        private const int ServerCallRetryLater = 2;
        private const int RetryDelayMs = 100;

        public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return 0;
        }

        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == ServerCallRetryLater)
            {
                return RetryDelayMs;
            }

            return -1;
        }

        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return 2;
        }
    }

    private sealed class Registration : IDisposable
    {
        private IOleMessageFilter? _previous;
        private bool _disposed;

        public Registration(IOleMessageFilter? previous)
        {
            _previous = previous;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            CoRegisterMessageFilter(_previous, out _);
            _previous = null;
        }
    }

    private sealed class NoopDisposable : IDisposable
    {
        public static NoopDisposable Instance { get; } = new();

        public void Dispose()
        {
        }
    }
}
