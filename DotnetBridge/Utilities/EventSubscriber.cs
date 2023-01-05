using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Westwind.WebConnection
{
    /// <summary>
    /// FoxPro interop access to .NET events. Handles all events of a source object for subsequent retrieval by a FoxPro client.
    /// </summary>
    /// <remarks>For a FoxPro program to be notified of events, it should use `wwDotNetBridge.InvokeMethodAsync` to call <see cref="WaitForEvent"/>. When <see cref="WaitForEvent"/> asynchronously completes, the FoxPro program should handle the event it returns and then call <see cref="WaitForEvent"/> again to wait for the next event. The FoxPro class `EventSubscription`, which is returned by `SubscribeToEvents`, encapsulates this async wait loop.</remarks>
    public sealed class EventSubscriber : IDisposable
    {
        private readonly object _source;
        private readonly List<DelegateInfo> _eventHandlers = new List<DelegateInfo>();
        private readonly ConcurrentQueue<RaisedEvent> _raisedEvents = new ConcurrentQueue<RaisedEvent>();
        private TaskCompletionSource<RaisedEvent> _completion = new TaskCompletionSource<RaisedEvent>();

        public EventSubscriber(object source, String prefix = "", dynamic vfp = null)
        {
            // Indicates that initially the client is not waiting.
            _completion.SetResult(null);

            // For each event, adds a handler that calls QueueInteropEvent.
            _source = source;
            foreach (var ev in source.GetType().GetEvents())
            {
                // handler is a PRIVATE variable defined in EventSubscription.Setup().
                Boolean hasMethod = vfp?.Eval($"PEMSTATUS(m.handler, '{prefix}{ev.Name}', 5)");
                if (!hasMethod)
                    continue;

                var eventParams = ev.EventHandlerType.GetMethod("Invoke").GetParameters().Select(p => Expression.Parameter(p.ParameterType)).ToArray();
                var eventHandlerLambda = Expression.Lambda(ev.EventHandlerType,
                    Expression.Call(
                        instance: Expression.Constant(this),
                        method: typeof(EventSubscriber).GetMethod(nameof(QueueInteropEvent), BindingFlags.NonPublic | BindingFlags.Instance),
                        arg0: Expression.Constant(ev.Name),
                        arg1: Expression.NewArrayInit(typeof(object), eventParams.Select(p => Expression.Convert(p, typeof(object))))),
                    eventParams);
                var eventHandler = eventHandlerLambda.Compile();
                ev.AddEventHandler(source, eventHandler);
                _eventHandlers.Add(new DelegateInfo(eventHandler, ev));
            }
        }

        class DelegateInfo
        {
            public DelegateInfo(Delegate handler, EventInfo eventInfo)
            {
                Delegate = handler;
                EventInfo = eventInfo;
            }

            public Delegate Delegate { get; }
            public EventInfo EventInfo { get; }
        }

        public void Dispose()
        {
            foreach (var item in _eventHandlers)
                item.EventInfo.RemoveEventHandler(_source, item.Delegate);
            _completion.TrySetCanceled();
        }

        private void QueueInteropEvent(string name, object[] parameters)
        {
            var interopEvent = new RaisedEvent { Name = name, Params = parameters };
            if (!_completion.TrySetResult(interopEvent))
                _raisedEvents.Enqueue(interopEvent);
        }

        /// <summary>
        /// Waits until an event is raised, or returns immediately if a queued event is available.
        /// </summary>
        /// <returns>The next event, or null if this subscriber has been disposed.</returns>
        public RaisedEvent WaitForEvent()
        {
            if (_raisedEvents.TryDequeue(out var interopEvent)) return interopEvent;
            _completion = new TaskCompletionSource<RaisedEvent>();
            var task = _completion.Task;
            
            task.Wait();

            return task.IsCanceled ? null : task.Result;
        }
    }

    public class RaisedEvent
    {
        public string Name { get; internal set; }
        public object[] Params { get; internal set; }
    }
}
