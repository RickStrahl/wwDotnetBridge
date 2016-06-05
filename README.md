# wwDotnetBridge

#### .NET Interop made easy for Visual FoxPro 9

wwDotnetBridge is a small library designed to make it easy to call .NET components from Visual FoxPro. By providing an easy mechanism for loading .NET components and calling them without requiring explicit COM registration of .NET components, it's easy to add .NET functionality to your applications. Interact with core .NET framework components, access both free or commercial 3rd party libraries, or build and access your own .NET components from FoxPro.

wwDotnetBridge also provides a host of tools to make it possible to access .NET type features that FoxPro and COM do not natively support. For example, COM cannot access components with multiple constructors, cannot access Value, Static or Generic members and types. wwDotnetBridge can automatically convert some problem types, and provides wrapper types that allow access to most unsupported feature types. There's also a powerful ComArray class that makes it easy to interact and manipulate .NET arrays and collections. 

### Features at a glance
wwDotnetBridge provides the following enhancements over plain COM Interop:

* Access most .NET Components directly even those not marked [ComVisible]
* Requires no COM Registration for .NET Assemblies
* Call any .NET method and set and get properties
* Call any .NET method asynchronously and get called back on completion
* Instantiate types with parameterized constructors
* Call overloaded methods using standard method names
* Support for many natively unsupported .NET types and values
* Access Static members, Value/Struct types, Generics, Binary, Guids, DbNulls
* Automatically fix up problematic .NET Types on method returns
* Provides easy Array access with ComArray helper class
* ComValue class to store results and parameters in .NET
* ComValue can help with problem .NET type conversions
* ComArray class that wraps all returned Array and Collections
* ComArray allows easy creation, updating and managing of Enumerable types
* Multi-threading library built-in
* wwDotnetBridge also works with regular COM Interop (w/o runtime hosting)


### Online Documentation:
* [wwDotnetBridge Home Page](http://west-wind.com/wwDotnetBridge.aspx)
* [.NET COM Interop with wwDotnetBridge White Paper](http://west-wind.com/presentations/wwDotnetBridge/wwDotnetBridge.pdf)
* [Documentation](http://west-wind.com/webconnection/docs/?page=_24n1cfw3a.htm)
* [ChangeLog](https://github.com/RickStrahl/wwDotnetBridge/blob/master/Changelog.md)

### How it works
This library consists of 3 components (all provided in source):

* ClrHost.dll - Win32 Loader for the .NET Runtime
* wwDotnetBridge.dll  - .NET assembly Proxy and Helper
* wwDotnetBridge.prg - FoxPro front end to .NET Proxy

If you're using FoxPro you can simply use the wwDotnetBridge.prg class to load an assembly and fire away at it or access existing .NET components directly.

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
The example demonstrates a few simple features: Loading an assembly, creating a .NET type instance and then calling methods and accessing properties either directly or indirectly. For many methods and properties on .NET object references you can directly access the members, but some members are not directly callable via COM if there are overloads on a method, if there are Generics, enums or Value Types involved in the method or member access. It's best to always try direct access first and if that fails attempt to use indirect access to the wwDotnetBridge instance.

In that case you can use indirect referencing to access members with:

* loBridge.InvokeMethod(instance,"Method",parm1,parm2..,parm15)
* loBridge.GetProperty(instance,"Property")  
* loBridge.SetProperty(instance,"Property",valueToSet)
* GetPropertyEx(),SetPropertyEx,InvokeMethodEx() which supported nested names for the member

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

### Resources
There's much more functionality available. Please check out the documentation for more info.

* [wwDotnetBridge Home Page](http://west-wind.com/wwDotnetBridge.aspx)
* [Documentation](http://west-wind.com/webconnection/wwClient_docs?page=_24n1cfw3a.htm)


## Project Sponsors
The following people/organizations have provided sponsorship to this project by way of direct donations or for paid development as part of a development project using these tools:

### West Wind Technologies
wwDotnetBridge was originally developed for [West Wind Client Tools](http://west-wind.com/webconnection) and [West Wind Web Connection](http://west-wind.com/wconnect), which continue to include a slightly modified version of wwDotnetBridge. West Wind Technologies has kindly open sourced wwDotnetBridge to extend the reach of FoxPro just a bit longer by allowing easy integration with .NET and allowing more people to access this useful functionality.

wwDotnetBridge updates are initially developed for both of the commercial products with any changes merged into this project when changes are made. The commercial versions also include a few add-on features used by the products. If you want a fully supported version of wwDotnetBridge or would like to sponsor further development efforts on wwDotnetBridge, you can show your support by optionally purchasing a license for either of these products.

* [West Wind Web Connection](http://west-wind.com/webconnection)
* [West Wind Internet and Client Tools](http://west-wind.com/WestwindClientTools.aspx)


#### Craig Tucker - Alabama Software
Craig offered early support and feedback for this project and billed project time for a number of additions to the library as part of a larger project.


#### Bill Suthman - Monosynth
Bill provided a sizable donation to the project and valuable feedback for a host of improvements and bug fixes.

#### Sunil Rjamara  - WeatherTrend
Sunil required a number of custom integrations into their FoxPro product that resulted in discovery of a number of edge cases that ended up getting integrated into wwDotnetBridge. WeatherTrend kindly donated a chunk of billable time to adding a handful of these small features.

Want to sponsor this project, need customization or want make a donation to show your support? You can contact me directly at rstrahl@west-wind.com or you can also make a donation online via PayPal.

* [Make a donation for wwDotnetBridge using PayPal](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=3CY6HGRTHSV5Y)
* [Make a donation for wwDotnetBridge using our Web Store](http://store.west-wind.com/product/donation)
* [Purchase a license for West Wind Internet and Client Tools](http://store.west-wind.com/product/wwclient50/)


## License
This library is licensed under **MIT license** terms:

> Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

<small>&copy; 2012-2016 Rick Strahl, West Wind Technologies</small>

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

## NO WARRANTY
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 
 