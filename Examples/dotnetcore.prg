* Requires .NET Core 3.0 or later
CLEAR
SET MEMOWIDTH TO 255

do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CREATEOBJECT("wwDotnetCoreBridge")
? loBridge.GetDotnetversion()

? loBridge.GetStaticProperty("System.Runtime.InteropServices.RuntimeInformation","FrameworkDescription")
loArchitecture  = loBridge.GetStaticProperty("System.Runtime.InteropServices.RuntimeInformation","OSArchitecture")
? loBridge.Getenumstring("System.Runtime.InteropServices.Architecture",loArchitecture) 

? "Network available " + TRANSFORM(loBridge.InvokeStaticMethod("System.Net.NetworkInformation.NetworkInterface",;
                              "GetIsNetworkAvailable"))
