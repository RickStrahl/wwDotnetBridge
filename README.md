# wwDotnetBridge
### .NET Interop made easy for Visual FoxPro 9.0

wwDotnetBridge is a small library designed to make it easy to **call .NET components from Visual FoxPro**. By providing an easy mechanism for loading .NET components and calling them **without requiring explicit COM registration of .NET components**, it's easy to add .NET functionality to your applications. Interact with core .NET framework components, access  system and free or commercial 3rd party libraries, or build and access your own .NET components and call them from FoxPro all without requiring COM registration.

wwDotnetBridge also provides a host of support features to make it possible to access .NET type features that FoxPro and COM alone do not natively support. For example, native COM interop cannot access components with multiple constructors, value types, .NET Generic types or static methods or properties.

wwDotnetBridge can automatically convert problem types via proxy methods like `InvokeMethod()`, `GetProperty()` and `SetProperty()`, or static versions like `InvokeStaticMethod()`, `GetStaticProperty()`, `SetStaticProperty()`. To help with problematic type access there's a powerful `ComArray` class that makes it easy to interact and manipulate .NET arrays, lists, collections and dictionaries from FoxPro, and a `ComValue` class that lets you assign, access and pass .NET values without ever passing the native .NET value into FoxPro, which allows you to access types that COM or FoxPro simply cannot directly access.

## Features at a glance

* Registrationless access to most .NET components
* Explicit loading of .NET assemblies from disk or the GAC
* Support for .NET Framework Runtime (4.72 or later)
* Support for .NET Core Runtime (via `wwDotnetCoreBridge`)
* Access to most .NET components from FoxPro
* Simple direct access to members of classes for any COM/FoxPro supported types
* Problem types and features are supported via helpers to:
    * Access to Value Types
    * Access to static methods and values
    * Access to enumerated values
    * Access to Generic types
    * Support for Async Task methods
    * Support for .NET Event Interfaces
    * Support for calling any .NET method asynchronously
    * [ComArray](https://webconnection.west-wind.com/docs/_6gg0mvpeo.htm) wrapper for .NET Arrays, Lists and Dictionaries
    * [ComValue](https://webconnection.west-wind.com/docs/_3481232sd.htm) for accessing COM/FoxPro incompatible types
    * Auto-Conversion of many problem .NET types to and from FoxPro
* DataSet conversions to and from XmlAdapter (and from XmlAdapter to cursors)
* ToJson() and ToXml() for .NET objects  (Commercial Version only)

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


## Getting Started
Typical steps for working with wwDotnetBridge are:

* Initialize the .NET Runtime during Startup
* Instantiate a wwDotnetBridge instance
* Use `.CreateInstance()` to create .NET Objects
* Use direct property and method access if supported by COM and FoxPro
* Use `.GetProperty()`, `.SetProperty()` and `.InvokeMethod()`    
to indirectly access methods and members that COM or FoxPro don't support
* Call Static methods and properties usoing static versions of the intrinsic functions

### Getting Started
The first thing that you need to do is *start up* wwDotnetBridge and you can do that by creating an instance with this code:

```foxpro
DO wwDotnetBridge                && Load the Library
loBridge = GetwwDotnetBridge()   && Get an instance of the FoxPro Proxy

*** Optionally check versions of library, .NET and Windows
? loBridge.GetDotnetVersion()
```

`GetwwDotnetBridge()` is a FoxPro helper function that creates *a cached instance* of the `wwDotnetBridge` object, that is reused and stored as a `PUBLIC` variable. This makes it very fast to get an instance of `wwDotnetBridge` after first load and it's the recommended way to get an instance, although `CREATEOBJECT("wwDotnetBridge")` also works albeit a little slower.

> First load of `wwDotnetBridge` can take a second or two as the .NET Runtime is loaded into the your FoxPro host process, but subsequent loads are nearly instant.

> #### @icon-warning  Unable to load CLR Instance Errors
> If you get an  <b>Unable to CLR Instance</b> error when creating an instance of wwDotnetBridge, please see <%= TopicLink([Unable to load CLR Instance],[_3RF12JTMA]) %> for more info.

Here's a rudimentary example that shows how wwDotnetBridge works:

```foxpro
*** Create or get cached instance of wwdotnetbridge
LOCAL loBridge as wwDotnetBridge, loHttp
loBridge = GetwwDotnetBridge()

*** Create a built-in .NET class and run a method
loHttp = loBridge.CreateInstance("System.Net.WebClient")
loHttp.DownloadFile("http://west-wind.com/files/MarkdownMonsterSetup.exe",
                    "MarkdownMonsterSetup.exe")
DO wwUtils
GoUrl(FULLPATH("MarkdownMonsterSetup.exe"))  && run it

*** Load a custom .NET assembly
loBridge.LoadAssembly("CustomDotnet.dll")

*** Access a .NET component from the new assembly
loItem = loBridge.CreateInstance("Custom.Item")

*** Access properties and methods directly
? loItem.Sku
loItem.Sku = "NewSku"
lnTotal = loItem.CalculateTotal()

*** Access non-accessible properties and methods indirectly
lnFlagValue = loBridge.GetProperty(loItem,"Flag")
lnFlagValue = loBridge.SetProperty(loItem,"Flag",5) 
loBridge.InvokeMethod(loItem,"PassFlagValue",lnFlagValue)
```

Note that not all properties and methods can be accessed directly as shown on the first example, but some properties and methods require implicit activation as in the 'Flag' example requiring `GetProperty()`, `SetProperty()` or `InvokeMethod()` to indirectly access object members. 

> If direct access fails, always try the indirect methods.

### Accessing Arrays, Collections and Dictionaries
A very common use case involves accessing collection types in .NET which can't be directly translated into FoxPro due to FoxPro and COM's very limited Array support. To help with this `wwDotnetBridge` provides a helper [ComArray class](VFPS://Topic/_6GG0MVPEO) that wraps .NET collection types and exposes a proxy interface to capture and interact with these types.

```foxpro
loBridge = GetwwDotnetBridge()    && instance
loBridge.LoadAssembly("wwDotnetBridgeDemos.dll")

*** Create an class Instance
loPerson = loBridge.CreateInstance("wwDotnetBridgeDemos.Person")

*** Access simple Properties
? "*** Simple Properties:" 
? loPerson.Name
? loPerson.Entered
?

*** Addresses is an Array use GetProperty() and return as a ComArray instance
loAddresses = loBridge.GetProperty(loPerson, "Addresses")  

? TRANSFORM(loAddresses.Count) + " Addresses"     && Number of items in array (2)

? "*** First Address"
loAddress = loAddresses.Item(0)
? "Street: " + loAddress.Street
? 

? "*** All Addresses"     && .NET Collections are 0 based!
FOR lnX = 0 TO loAddresses.Count-1
	loAddress = loAddresses.Item(lnX)
	? loAddress.ToString()  && writes out full address
	?
ENDFOR

? "*** Add a new Address to the array"
* loNewAddress = loBridge.CreateInstance("wwDotnetBridgeDemos.Address")
loNewAddress = loAddresses.CreateItem()
loNewAddress.Street = "122 Newfound Landing"
loNewAddress.City = "NewFoundLanding"
loAddresses.Add(loNewAddress)

? loAddresses.Count                 && 3 addresses now
loAddress = loAddresses.Item(2)     && retrieve the added address (0 based)
```

## Documentation
* [Home Page](https://west-wind.com/wwDotnetBridge.aspx)
* [API Documentation](https://client-tools.west-wind.com/docs/_24n1cfw3a.htm)
* [White Paper](http://west-wind.com/presentations/wwdotnetbridge/wwDotnetBridge.pdf)
* [Change Log](https://github.com/RickStrahl/wwDotnetBridge/blob/master/Changelog.md)

## How it works
This library consists of 3 components (all provided in source):

* ClrHost.dll - Win32 Loader for the .NET Runtime
* wwDotnetBridge.dll  - .NET assembly Proxy and Helper
* wwDotnetBridge.prg - FoxPro front end to .NET Proxy

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

You can also call any .NET method asynchronously and get called back when the method has completed executing:

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
DEFINE CLASS HttpCallback as Custom


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

## Events

.NET delegates and events are not directly supported. If the .NET object is available as a registered COM object, the COM events which translate to native FoxPro events; COM event binding is separate from wwDotNetBridge.

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
