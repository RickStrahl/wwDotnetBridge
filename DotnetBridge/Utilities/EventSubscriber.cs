#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace Westwind.WebConnection
{
    /// <summary>
    /// Subscribes to all events for which a handler object has corresponding methods.
    /// </summary>
    public sealed class EventSubscriber : IDisposable
    {
        private readonly object _eventSource;
        private readonly object _handler;
        private readonly bool _post;
        private readonly List<(EventInfo, Delegate)> _eventDelegates = [];

        private static readonly MethodInfo invokeMethod = typeof(EventSubscriber).GetMethod(nameof(InvokeMethod), BindingFlags.NonPublic | BindingFlags.Instance);

        public EventSubscriber(object eventSource, object handler, string prefix, bool post, dynamic vfp)
        {
            _eventSource = eventSource;
            _handler = handler;
            _post = post;
            var instanceExpression = Expression.Constant(this);
            var handlerExpression = Expression.Constant(handler);

            foreach (var eventInfo in eventSource.GetType().GetEvents())
            {
                string methodName = prefix + eventInfo.Name;
                bool hasMethod = vfp.Eval($"PEMSTATUS(m.wwDotnetBridgeEventHandler, '{methodName}', 5)");
                if (!hasMethod)
                    continue;

                var eventHandlerType = eventInfo.EventHandlerType;
                var paramExpressions = eventHandlerType.GetMethod("Invoke").GetParameters().Select(p => Expression.Parameter(p.ParameterType, p.Name)).ToArray();
                var arguments = paramExpressions.Select(p => Expression.Convert(p, typeof(object))).ToArray();
                var callExpression = Expression.Call(instanceExpression, invokeMethod, handlerExpression, Expression.Constant(methodName), Expression.NewArrayInit(typeof(object), arguments));
                var eventDelegate = Expression.Lambda(eventHandlerType, callExpression, paramExpressions).Compile();
                eventInfo.AddEventHandler(eventSource, eventDelegate);
                _eventDelegates.Add((eventInfo, eventDelegate));
            }
        }

        public void Dispose()
        {
            foreach ((var eventInfo, var eventDelegate) in _eventDelegates)
                eventInfo.RemoveEventHandler(_eventSource, eventDelegate);
            Marshal.FinalReleaseComObject(_handler);
        }

        private void InvokeMethod(object handler, string methodName, object[] arguments)
        {
            if (_post || Thread.CurrentThread != wwDotNetBridge._mainThread)
                wwDotNetBridge._synchronizationContext.Post(_ => handler.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, handler, arguments), null);
            else
                handler.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, handler, arguments);
        }
    }
}
