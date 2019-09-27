# wwDotnetBridge Change Log

<small>[Latest Online Documentation](http://west-wind.com/webconnection/docs/_24n1cfw3a.htm)</small>

### Version 7.09

* **Add improved Support for Task**  
When using intrinsic methods to get/set/call values `Task<T>` is now fixed up with `ConfigureAwait()` to ensure you can access the `Result` property of the Task to wait for completion.

* **ComValue.SetValueFromGenericInstance()**  
New helper method that allows assignment of a ComValue from a Generic type value. Since Generics aren't natively supported this method is one way to pass or receive Generic values from and to .NET.

* **ComGuid Deprecated**  
Moved all code that used to convert to and from `ComGuid` to work with `ComValue` for consistency. This might be a breaking change if you were working with .NET Guids previously. Instead of retrieving `GuidString` call `GetGuid()` instead, and when passing use `CreateValue()` and `SetGuid()`, `NewGuid()` to assign guids.

* **Fix `ComArray.SetItem()` for Property Fixups**  
Fixed `SetItem()` to use the same internal `FixupParameters` logic that understands common conversions from FoxPro to .NET. Previously this wasn't hooked up. Note `AddItem()` worked correctly.

* **First cut at wwDotnetCoreBridge to host and access .NET Core**  
There's now a `wwDotnetBridgeCore` class that can be used to host .NET Core and access .NET Core components. It works with .NET Core 3.0 and requires that .NET Core 3.0 is installed.  This is still experimental but it does work. To work you'll need the **32 bit .NET Runtime** installed. There are limitations currently - you can't unload the runtime and re-run it (including clearing all) without restarting. But it works fairly.

### Version 6.22

* **Added Event Subscription Support**  
You can now use `.SubscribeToEvents()` to subscribe to events of an object and pass a FoxPro object that implements the event methods using an specified (`On` by default) prefix. Thanks to Edward Brey for his work on this!

* **Add CLR 4.0 Hosting and better Loader Error Handling**  
Added new CLR hosting logic to the C++ connector which provides better support for .NET 4 hosting. Also added additional error handling for the CLR loader process so more error information may be available. Thanks to Edward Brey for his work on this!

* **ComValue.SetValueFromGenericInstance()**  
Added a helper to the ComValue object to allow creating of Generic Values. Since generic values and objects are not allowed over COM this is one option to assign and retrieve generic values.


### Version 6.19

* **wwDotnetBridge.dll is now compiled with .NET 4.5**  
This version of wwDotnetBridge is compiled for and requires .NET 4.5 or later. 4.5 offers a number of performance and feature enhancements that are required to load .NET 4.5 and later components. The vast majority of Windows machines come pre-installed with a version of .NET 4.5 or later so this version is the now the default.

* **Better Error information for wwDotnetBridge Load Errors**  
wwDotnetBridge will now report better error information on load failures, when failing to load the runtime or initial wwdotnetbridge instance.

### Version 6.15
*June 27th, 2017*

* **InitializeDotnetVersion()**  
An explicit function to initialize the .NET Runtime for the entire application. Although this function doesn't really do anything but create an instance of wwDotnetBridge with a specific version, the function makes it clear that this initializes the .NET version for the entire application.

* **Add explicit support for TLS 1.2**  
The DLL now explicitly allows TLS 1.2 without setting config file settings. This was causing some issues with http connections to servers that explicitly require TLS 1.2.

* **.NET and C++ Solution Compiled on VS2017**   
Both the C# and C++ projects now compile under Visual Studio 2017. You'll need the VS 2017 XP C++/ATL tooling installed in VS 2017y in order to compile ClrHost.

### Version 6.0
*June 8th, 2016*

* **wwDotnetBridge.dll is now compiled with .NET 4.0**  
This version of wwDotnetBridge is compiled for and requires .NET 4.0. 4.0 is offers a number of performance and feature enhancements for developement as well as looser security policy when executing code - we've long been recommending to **always** use .NET 4.0. In this version .NET 4.0 is now the default and this assembly is compiled for .NET 4.0. .NET 4.0 is the default installation of .NET with Windows 7 and later and .NET 2.0 is no longer installed by default on Windows 10, so 4.0 is the most common version installed. This change should not break any code accessing V2 assemblies, but you will need to run with .NET 4.x or later installed.


* **InvokeMethodAsync() and InvokeStaticMethodAsync()**  
You can now invoke any .NET method asynchronously and optionally receive a callback when the async operation is complete. New method that allows invoking .NET methods on an instance asynchronously. You pass in callback object that receives `OnCompleted()` and `OnError()` callbacks to get notified of completion of the async operation.

* **wwDotnetBridge::GetIndexedProperty()**    
New method adds the ability to retrieve an indexed value from an IList based object like arrays and generic lists more easily.

* **ComValue::SetValueFromSystemConvert()**  
Method that allows wwDotnetBridge to set a value from the .NET `System.Convert` static class in a simpler way. System.Convert() will try to coerce a value to a specific type more aggressively so things like a string can be converted to a number for example.

* **ComValue::GetTypeName()**  
Provides an easy way to see what the actual type stored in the `.Value` structure is. Useful for debugging.

* **ComValue::SetUInt64 and SetUInt32 Methods**  
Added additional type conversions for UInt64 and UInt32 types.

* **Fix Static Property Read and Assignments to fix up Values**  
Fix issue where `SetStaticProperty()` and `GetStaticProperty()` would not properly fix up common .NET and FoxPro type conversions where known problems exist (long->int,DbNull->null etc.). Also caused problems on the new Static assignment methods on the ComValue object.

* **Add ComValue.ToString() Method**  
Add method that returns the ToString() method result from the stored .NET value. Avoids having to use InvokeMethod() to get the value's string representation explicitly since the value is most likely not accessible from FoxPro.

* **Update ClrHost Win32 Project to VS2015**  
Updated the project to the new version of Visual Studio to minimize requirements for the VC++ VS2010 compiler. VS2015 Community is free and can be used by anybody to compile this code.

* **Cleanup examples**  
Cleaned up a number of the examples to represent recent version changes.


### Version 5.72
*April 9th, 2015*

* **Address concerns for running on XP**<br/>
Change compiler target using Visual Studio 2013 runtimes for ClrHost.dll.
This fixes potential problems on Windows XP.

* **Add support for char parameters and results**<br/>
char parameters can now be set with ComValue::SetChar() by passing in
either a string value or number. You can also use loBridge.ConvertToDotnetValue()
to create the ComValue structure. char results are converted to string when
returned from .NET method calls.

* **ComValue SetSingle()**<br/>
Allows assigning a .NET Single value from a FoxPro number to be used the 
dynamically invoked methods.

* **ComValue.SetChar() with characters or numbers**
.NET char parameters can now be set with ComValue::SetChar() by passing in either a string value or number. You can also use loBridge.ConvertToDotnetValue(val,"char")to create the ComValue structure. char results are converted to string when returned from .NET method calls. 

* **ComValue.SetValueFromEnum()**<br/>
Allows you to set an Enum Value from a type and constant name and 
assign it to ComValue. Allows passing Enum parameters to .NET which
otherwise is not possible as enums are static values. 

* **ComValue.SetValueFromInvokeStaticMethod**<br/>
You can now set a ComValue from the result of a static method
invocation. Useful if the result type of a static method returns a 
result that is not accessible in FoxPro - like a struct or generic
object. You can then use GetProperty/GetPropertyEx to access
values of that object.

* **Fixed error handling to be more consistent**<br/>
Changed the error handling behavior of all method that capture
errors to return the base error message. Made consistent across
wwDotnetBridge calls.

* **Add support for nested member strings to GetProperty/SetProperty/InvokeMethod**<br/>
You can now use . syntax in string based property/method names for all indirect methods, negating the need to use methods like GetPropertyEx or SetPropertyEx. You can now specify member names like "Address.Street" or "Error.ToHtml" directly in the various base methods. Use this feature to get around intermediary types that FoxPro can't access (Value types, generics etc.).

* **Better Support for Reference Parameters**<br/>
You can now pass parameters by reference using the indirect InvokeMethod/InvokeStaticMethod methods by passing parameters using a ComValue structure. The .Value property that holds the inbound parameter also receives any changed values that are modified by the method call.

* **Fix Null Value Handling**<br/>
Fixed bug with NULL values passed to wwDotnetBridge calls. COM Interop changes Fox NULLs to DbNulls which failed. Indirect methods now translate DbNull values to raw .NET nulls when passed. You can still pass DbNull with ComValue.SetDbNull() if needed.

### Version 5.67
*Nov. 8, 2013*

* **ComArray::AddItem() Auto Type Conversion**<br/>
When you add items to an array, AddItem() now automatically
recognizes ComValue and ComArray structures and performs
auto conversions of unsupported FoxPro types (like Guids, Longs etc.)

* **ComArray::Item() AutoType Conversions**<br/>
When you retrieve items out of the Item() method, values are 
automatically fixed up with type conversions for unssupported
FoxPro types. Arrays come back as ComArray, unsupported types
are converted to ComValue and some auto conversions like Longs
happen automatically.

* **Allow for 24 parameters in InvokeMethod()**<br/>
Due to many, many support requests with Invokemethod requirements
for an enormous amount of parameters, we bumped the supported parameter 
count to the max 24 parameters which is the maximum VFP allows 
(26 - 2 for object and method). For more parameters yet you can 
still use the InvokeMethod_ParameterArray() method. This is crazy but there you have it.

* **ComValue::SetValueFromCreateInstance()**
This method allows you to capture output from object creation and convert
it to one of the special types supported by COM value for fixups. ComValue
allows values to stay in .NET so you can access the supported special types
indirectly.

* **ComArray::FromEnumerable()**<br/>
This method allows you to capture any .NET Enumerable into a COM Array.
This includes access from arrays, List<T>, IQueryAble<T> etc. and allows
you to manipulate the resulting structure just like an array. Very useful
as many .NET components use abstract IEnumerables rather than explicit
arrays.