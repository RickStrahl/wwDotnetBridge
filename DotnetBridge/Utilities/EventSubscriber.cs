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
    /// FoxPro interop access to .NET events. Calls "OnEvent" on the event handler for each event, passing the event name and array of event parameters.
    /// </summary>
    /// <remarks>For each event is raised by the source, calls a method on the handler object named "On" plus the event name.</remarks>
    public sealed class EventSubscriber : IDisposable
    {
        private readonly object _source;
        private readonly object _handler;
        private readonly List<Delegate> _eventHandlers = new List<Delegate>();

        public EventSubscriber(object source, object handler)
        {
            _source = source;
            _handler = handler;
            var sourceType = source.GetType();
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
        }

        private void QueueInteropEvent(string name, object[] parameters)
        {
            _handler.GetType().InvokeMember("OnEvent", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, _handler, new object[] {
                name,
                new ComArray { Instance = parameters.Select(p => wwDotNetBridge.FixupParameter(p)).ToArray() }
            });
        }
    }
}
