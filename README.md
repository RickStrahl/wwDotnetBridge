# wwDotnetBridge
#### .NET Interop made easy for Visual FoxPro 9.0

![Bridge](Bridge.jpg)

wwDotnetBridge allows you to access the vast majority of .NET components directly from FoxPro. It provides registrationless activation of .NET Components, and acts as a proxy into .NET that makes it possible to access features that native COM Interop does not support directly.

To work around COM limitations, wwDotnetBridge provides many improvements and work arounds, while still using the base layer of COM Interop for the inter-process communication. Everything that works with native COM Interop also works with wwDotnetBridge - it's the same technology after all -  but you get many more support features and automatic type translations to work around the limitations.

### Links

* [wwDotnetBridge Documentation](https://webconnection.west-wind.com/docs/_24n1cfw3a.htm)    
* [wwDotnetBridge White Paper](https://west-wind.com/wconnect/weblog/ShowEntry.blog?id=57032)  
* [West Wind Message Board Support](https://support.west-wind.com)  
* [wwDotnetBridge Files](Distribution/)


### Key Features

* **Registrationless access to most .NET Components**  
Unlike native COM Interop, you can instantiate and access .NET Components and static classes, without requiring those classes to be registered as COM objects. Objects are instantiated from within .NET, so you can access most .NET components by directly loading them from their DLL assembly. Both .NET Framework (`wwDotnetBridge`) and .NET Core (`wwDotnetCoreBridge`) are supported.

* **Instantiates and Interacts with .NET Objects via COM from within .NET**  
wwDotnetBridge is a .NET based component that **runs inside of .NET** and acts as an intermediary for activation, invocation and access operations. It creates **any** .NET instances from within .NET and returns those references using COM Interop. Once loaded you can use all features that COM supports directly: Property access and method calls etc. *as long the members accessed use types that are supported by COM*.

* **Support for Advanced .NET Features that COM Interop doesn't support**  
Unfortunately there are many .NET features that COM and FoxPro don't natively support directly: Anything related to .NET Generics, overloaded methods, value types, enums, various number types to name just a few. But because wwDotnetBridge runs inside of .NET, it provides automatic conversions and helpers to allow access to these features via intermediary Reflection operations. These helpers access the unsupported COM operations from inside of .NET and translate the results into COM and FoxPro compatible results that are returned into your FoxPro application.

* **Automatic Type Conversions**  
Because there are many incompatible types in .NET that don't have equivalents in COM or FoxPro, wwDotnetBridge performs many automatic type conversions. These make it easier to call methods or retrieve values from .NET by automatically converting compatible types. For example: decimals to double, long, byte to int, Guid to string etc. There are also wrapper classes like `ComArray` that wraps  .NET Arrays and Collections and provides a FoxPro friendly interface for navigating and updating collections, and `ComValue` which wraps incompatible .NET values and provides convenient methods to set and retrieve the value in a FoxPro friendly way and pass it to .NET methods or property assignments.

* **Support for Async Code Execution**  
A lot of modern .NET Code uses async functionality via `Task` based interfaces, and wwDotnetBridge includes a `InvokeTaskMethodAsyc()` helper that lets you call these async methods and receive results via Callbacks asynchronously. You can also run **any** .NET synchronous method and call it asynchronously using `InvokeMethodAsync()` using the same Callback mechanism.

### Getting Started
The first step in using wwDotnetBridge is to load it for the first time, which instantiates the .NET Runtime. We recommend that you do this somewhere in your application startup sequence so as to avoid any potential version ambiguities. Somewhere in the startup of your application call `InitializeDotnetVersion()`:

```foxpro
DO wwDotnetBridge               && Loads dependencies
InitializeDotnetVersion("V4")   && Loads .NET Runtime and caches it
```

Note that `InitializeDotnetVersion()` is *optional*. You can use `GetwwDotnetBridge()` or `CREATEOBJECT("wwDotnetBridge")`, but using `InitializeDotnetVersion()` ensures a reliable and predictable load of .NET **on startup**.

> #### Unable to load CLR Instance Errors
> If you get an  <b>Unable to CLR Instance</b> error when creating an instance of wwDotnetBridge, you probably need to unblock the wwdotnetbridge.dll or need to ensure that the wwdotnetbridge.dll and wwipstuff.dll are in your FoxPro path. Please see <%= TopicLink([Unable to load CLR Instance],[_3RF12JTMA]) %> for more info.

> #### Loading DLLs from Network Locations: Configuration required
> .NET components require explicit configuration in order to support remote loading from network locations. This is done by creating a configuration file for your application `yourapp.exe.config` or the VFP IDE `vfp9.exe.config`, in their respective startup folders. We recommend at minimum you use the following `.config` file settings:
> ```xml
> <?xml version="1.0"?>
> <configuration>
>   <runtime>
>       <loadFromRemoteSources enabled="true"/>
>   </runtime>
> </configuration>
> ```

### wwDotnetBridge Example
With the library loaded, you can retrieve an instance by calling the `GetwwDotnetBridge()` factory function which caches a loaded wwDotnetBridge instance and therefore is very fast to access.

Here's an example what of some of what you can then do:

```foxpro
*** Create or get cached instance of wwdotnetbridge
LOCAL loBridge as wwDotnetBridge
loBridge = GetwwDotnetBridge()

*** The first two are built-in .NET Framework functions so no assembly has to be loaded

*** Create a built-in .NET class and execute a method - this one downloads a file to disk
loHttp = loBridge.CreateInstance("System.Net.WebClient")
loHttp.DownloadFile("http://west-wind.com/files/MarkdownMonsterSetup.exe",
                    "MarkdownMonsterSetup.exe")

*** Format a string: Static method: Typename as string, method, parameters
? loBridge.InvokeStaticMethod("System.String","Format","Hello {0}. Time is: {1:t}",;
                              "Rick", DATETIME())
* Hello Rick. Time is: 2:45 PM                              

*** Now load a third party Assembly - assemblies load their own dependencies!
? loBridge.LoadAssembly("wwDotnetBridgeDemos.dll")

*** Create a class Instance - naming is: namespace.class
loPerson = loBridge.CreateInstance("wwDotnetBridgeDemos.Person")

*** Access simple Properties - plain COM
? loPerson.Name
? loPerson.Company
? loPerson.Entered

*** Call simple method - plain COM
? loPerson.ToString()
? loPerson.AddAddress("1 Main","Fairville","CA","12345")  && 2 Addresses now

*** Access an Array/Collection of Objects and iterate over the list
*** Arrays/Collections/Dictionaries are not easily accessible via COM
*** wwDotnetBridge returns a `ComArray` instance
loAddresses = loBridge.GetProperty(loPerson,"Addresses");

*** Access ComArray.Count list count
lnCount = loAddresses.Count   && 2 addresses

*** Access the first item - 0 based list
loAddress = loAddress.Item(0);
? loAddress.Street
? loAddress.ToString()

*** Add another item to the array
* loNewAddress = loBridge.CreateInstance("wwDotnetBridgeDemos.Address")
loNewAddress = loAddresses.CreateItem()
loNewAddress.Street = "122 Newfound Landing"
loNewAddress.City = "NewFoundLanding"
loAddresses.Add(loNewAddress)

? TRANSFORM(loAddresses.Count) + " Addresses"  && 3

*** Iterate through the entire list (3 items now): Remember 0 based!
FOR lnX = 0 to loAddresses.Count -1 
    loAddress = loAddresses.Item(lnX)
    ? loAddress.ToString()
    ? 
ENDFOR
```

All interactions occur over COM so any object instances are COM objects with typical .NET COM behavior (no Intellisense, COM style errors). Any properties and methods that use standard types can be **directly accessed** via their normal COM property and method names. Any members or methods that use types that are incompatible with COM (Value types, Generics, Long, Decimal, Guid etc.) have to use the indirect access methods  `GetProperty()`, `SetProperty()` or `InvokeMethod()` for access.

> If direct access fails for whatever reason, always try the indirect methods.

For much more detailed wwDotnetBridge and .NET Interop information you can also check out the white paper:

* <a href="https://west-wind.com/wconnect/weblog/ShowEntry.blog?id=57032" target="top">wwDotnetBridge White Paper</a>.


> #### wwDotnetBridge and .NET Versions
> There are two versions of wwDotnetBridge, one for .NET Framework (2.0 - 4.8) and one for .NET Core (.NET Core 6.0+). 
> 
> Supported Platforms are:
>
> * .NET 4.7.2 Runtime or later  <small>*(.NET Framework - wwDotnetBridge)*</small>
> * .NET Core 6.0 Runtime and later <small>*(32 bit .NET Core - wwDotnetCoreBridge)*</small>
> * Windows 7 and newer
> * Windows Server 2012 and newer

## Documentation
You can find class and support documentation for wwDotnetBridge here:

* [wwDotnetBridge Documentation](https://webconnection.west-wind.com/docs/_24n1cfw3a.htm)
* [wwDotnetBridge Forum on West Wind Message Board](https://support.west-wind.com)


## How it works
This library consists of 3 components (all provided in source):

* ClrHost.dll - Win32 Loader for the .NET Runtime
* wwDotnetBridge.dll  - .NET assembly Proxy and Helper
* wwDotnetBridge.prg - FoxPro front end class to .NET Proxy

> #### Make sure DLLs can be found!
> Make sure CrlHost.dll (or wwIpstuff.dll for [commercial West Wind tools](https://west-wind.com/WestwindClientTools.aspx)) and wwDotnetBridge are accessible via the FoxPro path. Ideally you'll want to have these DLLs in your current executing path of the application - typically the root folder of the application.

If you're using FoxPro you can simply use the `wwDotnetBridge.prg` class to load an assembly and fire away at it or access existing .NET components directly.

This example loads the OpenPop .NET library and accesses a POP3 mailbox to display messages waiting:

```foxpro
*** Load library and initialize wwDotnetBridge
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge")

*** Load an assembly from disk
loBridge.LoadAssembly("bin\OpenPop.dll")

*** Create an instance of a class - note: No COM registration
loPop = loBridge.CreateInstance("OpenPop.Pop3.Pop3Client")

*** This won't work due to overloads
* loPop.Connect("mail.server.net",587,.f.)

*** So, call indirectly instead
? loBridge.InvokeMethod(loPop,"Connect","mail.server.net",110,.f.)

*** Most methods/members do work directly
? loPop.Authenticate("jb007","seekrit")

lnCount =  loPop.GetMessageCount()
? StringFormat("{0} Messages",lnCount)

*** NOTE: OpenPop is 1 based because pop3 is 1 based!
** show last messages
FOR lnX = lnCount TO 1 STEP -1
   loHeader = loPop.GetMessageHeaders(lnx)
   ? loHeader.From.DisplayName
   ? "  " + loHeader.Subject
   ?
   IF lnX < lnCount - 10
      EXIT
   ENDIF
ENDFOR
```
The example demonstrates a few simple features: Loading an assembly, creating a .NET type instance and then calling methods and accessing properties either directly or indirectly. You can call Static methods and access static members.

For many methods and properties on .NET object references you can directly access the members, but some members are not directly callable via COM if there are overloads on a method, if there are Generics, enums or Value Types involved in the method or member access. It's best to always try direct access first and if that fails attempt to use indirect access to the wwDotnetBridge instance.

In that case you can use indirect referencing to access members with:

* loBridge.InvokeMethod(instance,"Method",parm1,parm2..,parm15)
* loBridge.GetProperty(instance,"Property")  
* loBridge.SetProperty(instance,"Property",valueToSet)

Nested hierarchies can also be accessed for child properties or arrays:

* loBridge.GetProperty(instance,"Property.SubProperty")
* loBridge.GetProperty(instance,"Property.ArrayProp[0]") 

These methods internally use Reflection to call .NET code, but because they run inside of .NET they can do many things that native COM interop cannot due to the limitations for type marshalling over COM and the incompatibilities of the FoxPro type system.

Here's another example that demonstrates a few more features including calling static methods and setting enum values and accessing collections using enumerators.

```foxpro
*** Load library
DO wwDotNetBridge

*** Create instance of wwDotnetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")

loStore = loBridge.CreateInstance("System.Security.Cryptography.X509Certificates.X509Store")

? loBridge.cErrorMsg

*** Grab a static Enum value
leReadOnly = loBridge.GetEnumvalue("System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly")

*** Use the enum value
loStore.Open(leReadOnly)   && 0 - if value is known

*** Returns a .NET Collection of store items
laCertificates = loStore.Certificates

*** Collections don't work over regular COM Interop
*** so use indirect access
lnCount = loBridge.GetProperty(laCertificates,"Count")

*** Loop through Certificates
FOR lnX = 0 TO lnCount -1
	*** Access collection item indirectly using extended syntax
	*** that supports nested objects and array/collection [] brackets
	LOCAL loCertificate as System.Security.Cryptography.X509Certificates.X509Certificate2	
	loCertificate = loBridge.GetPropertyEx(loStore,"Certificates[" + TRANSFORM(lnX) + "]")
			
	IF !ISNULL(loCertificate)
		? loCertificate.FriendlyName
		? loCertificate.SerialNumber
		? loCertificate.GetName()
		*? loBridge.GetPropertyEx(loCertificate,"IssuerName.Name")
	ENDIF
ENDFOR
```


### Async Task Methods
`async` `await` using `Task` based operations are quite common in .NET. Task is similar to a JavaScript promise and basically encapsulates an operation that may or may not be complete yet. While Task based interfaces can use the `.Result` property to wait and retrieve results, this is potentially dangerous as UI events can cause the UI event loop to hang. Rather async Task methods should always be handled asynchronously.

You can also invoke `async/away` or `Task` based async methods in .NET via the `InvokeTaskMethodAsync()` method. Similar to the async approach above you need to provide a callback object. This works in a similar way that the arbitrary async operation works with `OnCompleted()` and `OnError()` handlers in a Callback object that are called when the async operation completes.

The following example uses `WebClient.DownloadDataTaskAsync()` which is a task based interface to download an image and displays it locally using asynchronous execution.

The C# signature that we are interested in for this method  is:

```cs
public async Task<byte[]> DownloadDataTaskAsync(string url)
```

The method asynchronously calls the URL in question and then returns a byte array (Blob response in FoxPro) when the async call completes.

In FoxPro with wwDotnetBridge we can run this code as follows:

```foxpro
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = GetwwDotnetBridge()

*** Create an instance of the callback object that gets called when done
loHttpCallback = CREATEOBJECT("HttpCallback")

lcImageUrl = "https://markdownmonster.west-wind.com/Images/MarkdownMonsterLogo.jpg"
loWebClient = loBridge.CreateInstance("System.Net.WebClient")

*** Pass the callback object, the object instance, method name, and any parameters
loBridge.InvokeTaskMethodAsync(loHttpCallback, loWebClient, "DownloadDataTaskAsync" ,lcImageUrl)

? "Making HTTP Call..."


* Callback Class - either OnCompleted or OnError is called
DEFINE CLASS HttpCallback as AsyncCallbackEvents

* Returns binary data
FUNCTION OnCompleted(lvImageBytes,lcMethod)

? "Got bytes..." + TRANSFORM(LEN(lvImageBytes))

lcOutputFIle = "c:\temp\test.jpg"
STRTOFILE(lvImageBytes, lcOutputFile)

DO wwutils && only to display the file
ShellExecute(lcOutputFile)

? "*** Done!"

ENDFUNC
*   OnCompleted

FUNCTION OnError(lcMessage, loException,lcMethod)
? lcMessage
ENDFUNC
*   OnError 

ENDDEFINE
```

The code works by way of a Callback object that is invoked when the async call completes. The operation is out of band similar to an event, and called **when the result is available**. The callback object is a separate object that is passed when calling the method, and the method then calls the `OnCompleted()` or `OnError()` handler on this callback object.

Async code is always a little more complex to write than sync code as the call and the result are not handled in a linear fashion. Instead we have the `OnCompleted()` and `OnError()` methods called separately from the mainline code. This may require some extra work to sync up the result handling with the application logic.

> #### Async Recommendations
> * Pass any state that you need into the Callback - as properties or on global objects/vars
> * Keep async callback code short
> * Minimize state changes in the callbacks

### Make any .NET Method call Async
Similar to Task methods you can also turn **any .NET Method call into an async call** using the same callback mechanism described above. The following calls the non-async version of  `WebClient.DownloadData()`:

```cs
public byte[] DownloadData(string url)
```

Although the call to this method is not async in .NET, we can **call it asynchronously anyway** using the `InvokeMethodAsync()` method. The following should look familiar from the last example:

```foxpro
DO wwDotnetBridge

LOCAL loBridge as wwDotnetBridge
loBridge = GetwwDotnetBridge("V4")

loHttp = loBridge.CreateInstance("System.Net.WebClient")


*** Create a callback object - object has to be 'global'
*** so it's still around when the callback returns
*** Use Public or attach to long lived object like form or _Screen
PUBLIC loCallback
loCallBack = CREATEOBJECT("HttpCallback")


*** Make the async call - returns immediately
loBridge.InvokeMethodAsync(loCallback, loHttp, ;
                        "DownloadData",;
                        "http://west-wind.com/files/HelpBuilderSetup.exe")

? "Download has started... running in background."

RETURN


*** Callback class 
DEFINE CLASS HttpCallback as AsyncCallbackEvents


FUNCTION OnCompleted(lvResult, lcMethod)

? "Http Call completed"
? "Received: " + TRANS(LEN(lvResult),"999,999,999")

lcFile = ADDBS(SYS(2023)) + "HelpBuilderSetup.exe"

*** Write to file to temp folder
STRTOFILE(lvResult,lcFile)

*** Launch the downloaded installer
TRY
	*** Open the Zip file
	DO wwutils
	GoUrl(lcFile)
CATCH
ENDTRY

ENDFUNC

FUNCTION OnError(lcErrorMessage, loException, lcMethod)
? lcErrorMessage
ENDFUNC

ENDDEFINE
```

### Events
wwDotNetBridge supports an alternative approach to obtain events without COM registration. You can subscribe to all events of a .NET object by calling `wwDotNetBridge.SubscribeToEvents`. You pass the source object for the events and your handler object. The class for your handler object should have an `On...` function for each event.

This example creates a `System.Net.Mail.SmtpClient` object and handles its one event, `SendCompleted`:

```foxpro
LOCAL loSmtpClient, loSmtpHandler, loSmtpEventSubscription
loSmtpClient = loBridge.CreateInstance("System.Net.Mail.SmtpClient")
loSmtpHandler = CREATEOBJECT("MySmtpEventHandler")
loSmtpEventSubscription = loBridge.SubscribeToEvents(loSmtpClient, loSmtpHandler)
* Send email here

DEFINE CLASS MySmtpEventHandler as Custom
PROCEDURE OnSendCompleted(loSender, loEventArgs)
* Handle the event here
ENDPROC
ENDDEFINE
```

When no longer want to be notified of the events, call `Unsubscribe` on the subscription (`loSmtpEventSubscription` in this example).

## Project Sponsors
The following people/organizations have provided sponsorship to this project by way of direct donations or for paid development as part of a development project using these tools:

### West Wind Technologies
wwDotnetBridge was originally developed for [West Wind Client Tools](http://west-wind.com/webconnection) and [West Wind Web Connection](http://west-wind.com/wconnect), which continue to include a slightly modified version of wwDotnetBridge. West Wind Technologies has kindly open sourced wwDotnetBridge to extend the reach of FoxPro just a bit longer by allowing easy integration with .NET and allowing more people to access this useful functionality.

wwDotnetBridge updates are initially developed for both of the commercial products with any changes merged into this project when changes are made. The commercial versions also include a few add-on features used by the products such as an SMTP client, SFTP support, Encryption and Image management utilities through .NET wrappers. If you want a fully supported version of wwDotnetBridge or would like to sponsor further development efforts on wwDotnetBridge, you can show your support by purchasing a license for either of these products.

* [West Wind Web Connection](http://west-wind.com/webconnection)
* [West Wind Internet and Client Tools](http://west-wind.com/WestwindClientTools.aspx)


### Craig Tucker - Alabama Software
Craig offered early support and feedback for this project and billed project time for a number of additions to the library as part of a larger project.

### Bill Suthman - Monosynth
Bill provided a sizable donation to the project and valuable feedback for a host of improvements and bug fixes.

### Sunil Rjamara  - WeatherTrend
Sunil required a number of custom integrations into their FoxPro product that resulted in discovery of a number of edge cases that ended up getting integrated into wwDotnetBridge. WeatherTrend kindly donated a chunk of billable time to adding a handful of these small features.

### Want to be a Sponsor?
Want to sponsor this project, need customization or want make a donation to show your support? You can contact me directly at rstrahl@west-wind.com or you can also make a donation online via PayPal.

* [Sponsor this GitHub account](https://github.com/sponsors/RickStrahl)
* [Make a donation for wwDotnetBridge using PayPal](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=3CY6HGRTHSV5Y)
* [Make a donation for wwDotnetBridge using our Web Store](http://store.west-wind.com/product/donation)
* [Purchase a license for West Wind Internet and Client Tools](http://store.west-wind.com/product/wwclient70/)


## License
This library is licensed under **MIT license** terms:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

<small>&copy; 2012-2018 Rick Strahl, West Wind Technologies</small>

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

### NO WARRANTY
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

<small>&copy; 2012-2023 Rick Strahl, West Wind Technologies</small>
