using System;
using System.Runtime.InteropServices;

namespace Westwind.WebConnection
{
    /// <summary>
    /// Can be used to create an instance of wwDotnetBridge
    /// from the .NET Core Runtime Host which requires a static method
    /// </summary>
    public class wwDotnetBridgeFactory
    {

        [return: MarshalAs(UnmanagedType.IUnknown)]
        public static wwDotNetBridge CreatewwDotnetBridge()
        {
            return new wwDotNetBridge();
        }

        public static int CreatewwDotnetBridgeByRef([MarshalAs(UnmanagedType.IUnknown)] ref object instance)
        {
            try
            {
                instance = new wwDotNetBridge();
            }
            catch (Exception ex)
            {
                instance = null;
                return ex.HResult;
            }

            return 0;
        }

        /// <summary>
        /// Test method that returns an integer value of 10 for testing
        /// </summary>
        /// <returns></returns>
        public int Ping()
        {
            return 10;
        }
    }
}