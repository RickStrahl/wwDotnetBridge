using System;
using System.Runtime.InteropServices;

namespace Westwind.WebConnection
{
    /// <summary>
    /// Can be used to create an instance of wwDotnetBridge
    /// from the .NET Core Runtime Host which requires a static method
    /// </summary>
    public unsafe class wwDotnetBridgeFactory
    {

        [return: MarshalAs(UnmanagedType.IUnknown)]
        public static wwDotNetBridge CreatewwDotnetBridge()
        {
            return new wwDotNetBridge();
        }

        [return: MarshalAs(UnmanagedType.IUnknown)]
        public static wwDotNetBridge CreatewwDotnetBridge(string parm)
        {
            return new wwDotNetBridge();
        }

        public static int CreatewwDotnetBridgeByRef([MarshalAs(UnmanagedType.IDispatch)] ref object instance)
        {
            try
            {
                instance = new wwDotNetBridge();
            }
            catch
            {
                instance = null;
                return -1;
            }

            return 0;
        }

        /// <summary>
        /// Test method that returns an integer value of 10 for testing
        /// </summary>
        /// <returns></returns>
        public static int Ping(ref string msg)
        {
            StringUtils.LogString( "Incoming wwDotnetBridge creation.", @"c:\temp\log.log");
            msg = "Hello World";
            return 10;
        }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Test
    {

        public string Name { get; set; } = "Rick";
        public int Number { get; set; } = 1;
    }
}