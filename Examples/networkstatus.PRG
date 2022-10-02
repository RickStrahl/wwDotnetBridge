CLEAR

*** Demonstrates:
*** Accessing a static method in .NET
do wwDotNetBridge

LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")

llResult = loBridge.InvokeStaticMethod("System.Net.NetworkInformation.NetworkInterface",;
                              "GetIsNetworkAvailable")
                              
IF llResult
  ? "Network is available"
ELSE
  ? "Network unavailable"
ENDIF    
