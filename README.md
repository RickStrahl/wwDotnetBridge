# wwDotnetBridge
### 针对 Visual FoxPro 9 的更容易的 .NET 互操作

wwDotnetBridge是一个小型库，旨在使**从Visual FoxPro 调用.NET组件**变得容易。 通过提供一种简单的机制来加载.NET组件并在不需要.NET组件的明确COM注册的情况下调用它们，就可以轻松地将.NET功能添加到您的应用程序中。 与核心.NET框架组件交互，访问免费或商业第三方**库**，或者从FoxPro构建和访问您自己的.NET组件，而无需通过COM注册组件。

wwDotnetBridge还提供了大量工具，以便访问FoxPro和COM本身不支持的.NET类型功能。 例如，原生COM互操作无法访问具有多个构造函数的组件，无法访问Value，Static或Generic成员和类型。 wwDotnetBridge可以自动转换一些问题类型，并提供允许访问大多数不受支持的要素类型的包装类型。 还有一个功能强大的ComArray类，可以轻松交互和操作.NET数组和集合，还有一个ComValue类，可以让您分配和传递.NET值，而无需触及FoxPro代码。

#### wwDotnetBridge 和 .NET 版本
> 当前版本的wwDotnetBridge是为.NET 4.5编译的，适用于：
>
> * .NET 4.5 或更高
> * Windows 7 或更高
> * Windows Server 2008 R2 或更高
>
> 要获得Windows XP，Server 2003和2008 的支持，**您必须使用[版本6.0的wwDotnetBridge]（https://github.com/RickStrahl/wwDotnetBridge/releases/tag/v6.0）**，该版本的最新版本 用 **.NET 4.0** 编译，可以运行XP，Vista，Server2003 / 2008。 请注意，您可以使用新版本来加载.NET 1.1,2.0和4.0编译程序集。

## 从这里开始
Somewhere in the startup of your application call `InitializeDotnetVersion()`:

```foxpro
*** Load dependencies and add to Procedure stack
*** Make sure wwDotnetBridge.prg wwDotnetBridge.dll wwIPStuff.dll 
*** are in your FoxPro path
DO wwDotnetBridge
InitializeDotnetVersion("V4") 
```

This ensures that wwDotnetBridge loads with the specified **single version of the .NET Runtime** that your FoxPro application can load.

> #### @icon-warning  Unable to load CLR Instance Errors
> If you get an  <b>Unable to CLR Instance</b> error when creating an instance of wwDotnetBridge, you probably need to unblock the wwdotnetbridge.dll or need to ensure that the wwdotnetbridge.dll and wwipstuff.dll are in your FoxPro path. Please see [Unable to load CLR Instance](https://www.west-wind.com/webconnection/wwClient_docs/_3rf12jtma.htm) for more info.

Then when you need to utilize wwDotnetBridge call `GetwwDotnetBridge()` to get a cached instance and use it to access .NET components:

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

*** Access properties directly
? loItem.Sku
loItem.Sku = "NewSku"
lnTotal = loItem.CalculateTotal()

*** Access non-accessible properties and methods indirectly
lnFlagValue = loBridge.GetProperty(loItem,"Flag")
lnFlagValue = loBridge.SetProperty(loItem,"Flag",5) 
loBridge.InvokeMethod(loItem,"PassFlagValue",lnFlagValue)
```

## Features at a glance
wwDotnetBridge provides the following enhancements over plain COM Interop:

* Access most .NET Components directly even those not marked [ComVisible]
* Requires no COM Registration
* Call any .NET method and set and get properties
* Call any .NET method asynchronously and get called back on completion
* Instantiate types with parameterized constructors
* Call overloaded methods using standard method names
* Support for many natively unsupported .NET types and values
* Access Static members, Value/Struct types, Generics, Binary, Guids, DbNulls
* Automatically fix up problematic .NET Types on method returns
* Provides easy Array access with ComArray helper class
* ComValue class to store results and parameters in .NET
* ComValue helps with problem .NET type conversions
* ComArray allows easy creation, updating and managing of Enumerable types
* Multi-threading library built-in 
* wwDotnetBridge can also work with regular COM Interop (w/o runtime hosting)

## Documentation
* [Home Page](http://west-wind.com/wwDotnetBridge.aspx)
* [API Documentation](https://www.west-wind.com/webconnection/wwClient_docs/_24n1cfw3a.htm)
* [White Paper](http://west-wind.com/presentations/wwdotnetbridge/wwDotnetBridge.pdf)
* [Change Log](https://github.com/RickStrahl/wwDotnetBridge/blob/master/Changelog.md)

## How it works
This library consists of 3 components (all provided in source):

* ClrHost.dll - Win32 Loader for the .NET Runtime
* wwDotnetBridge.dll  - .NET assembly Proxy and Helper
* wwDotnetBridge.prg - FoxPro front end to .NET Proxy

> #### Make sure DLLs can be found!
> Make sure CrlHost.dll (or wwIpstuff.dll for [commercial West Wind tools](https://west-wind.com/WestwindClientTools.aspx)) and wwDotnetBridge are accessible via the FoxPro path. Ideally you'll want to have these DLLs in your current executing path of the application - typically the root folder of the application.

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

* [Make a donation for wwDotnetBridge using PayPal](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=3CY6HGRTHSV5Y)
* [Make a donation for wwDotnetBridge using our Web Store](http://store.west-wind.com/product/donation)
* [Purchase a license for West Wind Internet and Client Tools](http://store.west-wind.com/product/wwclient50/)


## License
This library is licensed under **MIT license** terms:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

<small>&copy; 2012-2018 Rick Strahl, West Wind Technologies</small>

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

### NO WARRANTY
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

<small>&copy; 2012-2018 Rick Strahl, West Wind Technologies</small>
