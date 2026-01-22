using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Westwind.WebConnection;

namespace wwDotnetBridge.Tests
{
    /// <summary>
    /// Tests FoxPro interop.
    /// </summary>
    [TestClass]
    public class EventSubscriberTests
    {
        private readonly wwDotNetBridge _bridge = new();
        private readonly TaskCompletionSource<object> _raisedCompletion = new();
        private bool _onNoParamsRaised;

        [TestInitialize]
        public void TestInitialize()
        {
            _bridge.SetSynchronizationContext(0);
        }

        [TestMethod]
        public Task EventSubscriber_RaiseImmediateEvent() => RaiseEvent(false);

        [TestMethod]
        public Task EventSubscriber_RaisePostedEvent() => RaiseEvent(true);

        private async Task RaiseEvent(bool post)
        {
            var loopback = new Loopback();
            var subscriber = new EventSubscriber(loopback, this, "On", post, this);
            loopback.Raise();
            if (post) // Conditional to verify that dispatching is not required for immediate events (even though it is harmless).
                _bridge.Dispatch();
            await _raisedCompletion.Task;
            Assert.ThrowsException<ArgumentException>(subscriber.Dispose); // Expect Marshal.FinalReleaseComObject to throw because our test handler is not a COM object.
        }

        public void OnNoParams()
        {
            _onNoParamsRaised = true;
        }

        public void OnTwoParams(string s, int i)
        {
            Assert.IsTrue(_onNoParamsRaised);
            Assert.IsTrue(s == "A" && i == 1);
            _raisedCompletion.SetResult(null);
        }

        public bool Eval(string _) => true;
    }

    public class Loopback
    {
        public event Action NoParams;
        public event Action<string, int> TwoParams;

        public void Raise()
        {
            NoParams();
            TwoParams("A", 1);
        }
    }
}
