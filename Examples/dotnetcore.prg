* Requires .NET Core 3.0 or later
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CREATEOBJECT("wwDotnetCoreBridge")

? loBridge.GetDotnetversion()

? "---"

*** Read the Runtime Version
? loBridge.GetStaticProperty("System.Runtime.InteropServices.RuntimeInformation","FrameworkDescription");

*** Call an arbitrary  API
? "Network available " + TRANSFORM(loBridge.InvokeStaticMethod("System.Net.NetworkInformation.NetworkInterface",;
                              "GetIsNetworkAvailable"))
                              
                              
*loBridge.Unload()
