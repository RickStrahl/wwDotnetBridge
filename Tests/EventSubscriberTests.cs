using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Westwind.WebConnection;

namespace Westwind.WebConnection.Tests
{
    /// <summary>
    /// Tests FoxPro interop.
    /// </summary>
    [TestClass]
    public class EventSubscriberTests
    {
        [TestMethod]
        public void EventSubscriber_WaitForPastEvents()
        {
            var loopback = new Loopback();
            var subscriber = new EventSubscriber(loopback);
            loopback.Raise();
            VerifyResults(subscriber);
        }

        [TestMethod]
        public void EventSubscriber_WaitForFutureEvents()
        {
            var loopback = new Loopback();
            var subscriber = new EventSubscriber(loopback);
            Task.Delay(1).ContinueWith(t => loopback.Raise());
            VerifyResults(subscriber);
        }

        static void VerifyResults(EventSubscriber subscriber)
        {
            var result = subscriber.WaitForEvent();
            Assert.IsTrue(result.Name == nameof(Loopback.NoParams) && result.Params.Length == 0);
            result = subscriber.WaitForEvent();
            Assert.IsTrue(result.Name == nameof(Loopback.TwoParams) && result.Params.Length == 2 && (string)result.Params[0] == "A" && (int)result.Params[1] == 1);
        }
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
