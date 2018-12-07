using System;
using System.Collections;
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
        private readonly List<Delegate> _eventHandlers = new List<Delegate>();
        private readonly ConcurrentQueue<RaisedEvent> _raisedEvents = new ConcurrentQueue<RaisedEvent>();

        /// <summary>
        /// Completed when an event is raised, or with null if the client is not waiting.
        /// </summary>
        private TaskCompletionSource<RaisedEvent> _completion = new TaskCompletionSource<RaisedEvent>();

        public EventSubscriber(object source)
        {
            // Indicates that initially the client is not waiting.
            _completion.SetResult(null);

            // For each event, adds a handler that calls QueueInteropEvent.
            _source = source;
            foreach (var ev in source.GetType().GetEvents()) {
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
                _eventHandlers.Add(eventHandler);
            }
        }

        public void Dispose()
        {
            var events = _source.GetType().GetEvents();
            for (int e = 0; e < events.Length; ++e)
                events[e].RemoveEventHandler(_source, _eventHandlers[e]);
            _completion.TrySetResult(null); // Releases the waiting client.
        }

        private void QueueInteropEvent(string name, object[] parameters)
        {
            var interopEvent = new RaisedEvent { Name = name, Params = new ArrayList(parameters) };
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
            return _completion.Task.Result;
        }
    }

    public class RaisedEvent
    {
        public string Name { get; internal set; }
        public ArrayList Params { get; internal set; } // An ArrayList is used here instead of object[] to work around difficulties with indexing into arrays from FoxPro.
    }
}
