*** Demonstrates:
*** Accessing a static method in .NET
do wwDotNetBridge

LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")

? loBridge.InvokeStaticMethod("System.Net.NetworkInformation.NetworkInterface",;
                              "GetIsNetworkAvailable")
