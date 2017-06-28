
# wwDotnetBridge Change Log

<small>[Latest Online Documentation](http://west-wind.com/webconnection/docs/_24n1cfw3a.htm)</small>

### Version 6.15
*June 27th, 2017*

* **InitializeDotnetVersion()**  
An explicit function to initialize the .NET Runtime for the entire application. Although this function doesn't really do anything but create an instance of wwDotnetBridge with a specific version, the function makes it clear that this initializes the .NET version for the entire application.

* **Add explicit support for TLS 1.2**  
The DLL now explicitly allows TLS 1.2 without setting config file settings. This was causing some issues with http connections to servers that explicitly require TLS 1.2.

* **.NET and C++ Solutiuon Compiled on VS2017**   
Both the C# and C++ projects now compile under Visual Studio 2017. You'll need the VS 2017 XP C++/ATL tooling installed in VS 2017y in order to compile ClrHost.

### Version 6.0
*June 8th, 2016*

* **wwDotnetBridge.dll is now compiled for .NET 4.0**  
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