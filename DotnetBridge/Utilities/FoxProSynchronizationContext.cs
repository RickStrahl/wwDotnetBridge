#nullable enable
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Westwind.WebConnection
{
    /// <summary>
    /// Synchronizes tasks with the FoxPro main thread.
    /// </summary>
    /// <remarks>
    /// When one or more tasks are ready to run, posts a Windows message to the main FoxPro window.
    /// Subclasses the window to receive the posted message and dispatch the tasks.
    /// Unlike FoxPro BINDEVENT, subclassing processes messages even when FoxPro pumps messages from a dispatched task (e.g. from a modal form).
    /// </remarks>
    internal sealed class FoxProSynchronizationContext : SynchronizationContext
    {
        private readonly IntPtr _hwnd;
        private readonly ConcurrentQueue<(SendOrPostCallback handler, object? state)> _postQueue = [];
        private readonly WndProcDelegate _wndProcDelegate;
        private readonly IntPtr _originalWndProc;
        private readonly uint _postMessageId;

        public FoxProSynchronizationContext(int hwnd)
        {
            _hwnd = (IntPtr)hwnd;
            _wndProcDelegate = WndProc; // Prevents the delegate from being garbage collected.
            _originalWndProc = SetWindowLongPtr(_hwnd, GWLP_WNDPROC, Marshal.GetFunctionPointerForDelegate(_wndProcDelegate));
            _postMessageId = RegisterWindowMessage("FoxProSynchronizationContextDispatch");
        }

        private IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
        {
            if (msg == _postMessageId)
            {
                Dispatch();
                return IntPtr.Zero;
            }

            return CallWindowProc(_originalWndProc, hWnd, msg, wParam, lParam);
        }

        /// <summary>
        /// Posts a message to indicate that there are posts ready to dispatch. Thread safe.
        /// </summary>
        public override void Post(SendOrPostCallback d, object? state)
        {
            _postQueue.Enqueue((d, state));

            if (!PostMessage(_hwnd, _postMessageId, IntPtr.Zero, IntPtr.Zero))
                throw new OutOfMemoryException("Failed to post dispatch message.");
        }

        /// <summary>
        /// Dispatches all queued send or post callbacks in the synchronization context.
        /// </summary>
        private void Dispatch()
        {
            while (_postQueue.TryDequeue(out var post))
            {
                try
                {
                    post.handler(post.state);
                }
                catch (Exception ex)
                {
                    // Reports the unhandled exception as an unobserved task exception.
                    // An event is raised by TaskScheduler.UnobservedTaskException after the task's finalizer runs.
                    var edi = ExceptionDispatchInfo.Capture(ex);
                    Task.Run(() => edi.Throw());
                }
            }
        }

        /// <summary>
        /// Starts a dispatch operation. Used by external code to dispatch queued callbacks.
        /// </summary>
        public override void OperationStarted() => Dispatch();

        private const int GWLP_WNDPROC = -4;

        private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll", EntryPoint = "SetWindowLongA")]
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        private static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll", CharSet = CharSet.Unicode, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        private static extern uint RegisterWindowMessage(string lpString);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    }
}
