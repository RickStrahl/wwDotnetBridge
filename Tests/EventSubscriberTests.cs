using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Westwind.WebConnection.Tests
{
    /// <summary>
    /// Test class that immediately raises events when <see cref="Raise"/> is called.
    /// </summary>
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
