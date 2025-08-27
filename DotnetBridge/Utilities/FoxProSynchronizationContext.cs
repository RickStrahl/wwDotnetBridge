#nullable enable
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Westwind.WebConnection
{
    /// <summary>
    /// Synchronizes tasks with the FoxPro main thread.
    /// </summary>
    /// <remarks>When one or more tasks are ready to run, posts a Windows message to the main FoxPro window.</remarks>
    sealed class FoxProSynchronizationContext(int hwnd, object taskWaiter) : SynchronizationContext
    {
        private readonly IntPtr _hwnd = (IntPtr)hwnd;
        private readonly ConcurrentQueue<(SendOrPostCallback handler, object? state)> _postQueue = [];
        private readonly dynamic _taskWaiter = taskWaiter;

        /// <summary>
        /// Gets the ID of the Windows message posted when a task is ready to run.
        /// </summary>
        public int PostMessageId { get; private set; } = RegisterWindowMessage("FoxProSynchronizationContextDispatch");

        public override void Post(SendOrPostCallback d, object? state)
        {
            _postQueue.Enqueue((d, state));
            PostMessage(_hwnd, PostMessageId, IntPtr.Zero, IntPtr.Zero);
        }

        /// <summary>
        /// Dispatches all queued send or post callbacks in the synchronization context. Called when a Windows message with ID <see cref="PostMessageId"/> is received.
        /// </summary>
        public void Dispatch()
        {
            while (_postQueue.TryDequeue(out var postedEvent))
                try { postedEvent.handler(postedEvent.state); } catch { }
        }

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll", CharSet = CharSet.Unicode, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        private static extern int RegisterWindowMessage(string lpString);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    }
}
