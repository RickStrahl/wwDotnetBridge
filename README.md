# wwDotnetBridge
### 针对 Visual FoxPro 9 的更容易的 .NET 互操作

翻译：xinjie 2018.05.10

**鸣谢：[龙岩耐思](http://www.fjlynice.com/)**

wwDotnetBridge 是一个小型库，旨在使**从 Visual FoxPro 调用 .NET 组件**变得容易。 通过提供一种简单的机制来加载 .NET 组件并在不需要 .NET 组件的明确 COM 注册的情况下调用它们，就可以轻松地将 .NET 功能添加到您的应用程序中。 与核心 .NET 框架组件交互，访问免费或商业第三方**库**，或者从 FoxPro 构建和访问您自己的 .NET 组件，而无需通过 COM 注册组件。

wwDotnetBridge 还提供了大量工具，以便访问 FoxPro 和 COM 本身不支持的 .NET 类型功能。 例如，原生 COM 互操作无法访问具有多个构造函数的组件，无法访问Value，Static或Generic成员和类型。 wwDotnetBridge 可以自动转换一些问题类型，并提供允许访问大多数不受支持的要素类型的包装类型。 还有一个功能强大的ComArray 类，可以轻松交互和操作 .NET 数组和集合，还有一个 ComValue 类，可以让您分配和传递 .NET 值，而无需触及 FoxPro 代码。

#### wwDotnetBridge 和 .NET 版本
> 当前版本的wwDotnetBridge是为.NET 4.5编译的，适用于：
>
> * .NET 4.5 或更高
> * Windows 7 或更高
> * Windows Server 2008 R2 或更高
>
> 要获得Windows XP，Server 2003和2008 的支持，**您必须使用[版本6.0的wwDotnetBridge](https://github.com/RickStrahl/wwDotnetBridge/releases/tag/v6.0)** ，该版本的最新版本 用 **.NET 4.0** 编译，可以运行XP，Vista，Server2003 / 2008。 请注意，您可以使用新版本来加载.NET 1.1,2.0和4.0编译程序集。

## 从这里开始
在你的应用程序启动的某个地方调用`InitializeDotnetVersion()`

```foxpro
*** 加载依赖关系并添加到过程堆栈
*** 确保 wwDotnetBridge.prg wwDotnetBridge.dll wwIPStuff.dll
*** 在您的 FoxPro 路径中
DO wwDotnetBridge
InitializeDotnetVersion("V4") 
```

这可以确保 wwDotnetBridge 加载您的 FoxPro 应用程序可以加载的指定**单一版本的.NET运行时**。

> #### @icon-warning  无法加载CLR实例错误
> 如果在创建wwDotnetBridge实例时收到<b>无法CLR实例</b>错误，则可能需要解除对 wwdotnetbridge.dll 的阻止或需要确保 wwdotnetbridge.dll 和wwipstuff.dll 位于您的 FoxPro 路径中。 请参阅[无法加载CLR实例](https://www.west-wind.com/webconnection/wwClient_docs/_3rf12jtma.htm)了解更多信息。

然后，当您需要利用 wwDotnetBridge 调用 `GetwwDotnetBridge()` 来获取缓存实例并使用它来访问 .NET 组件时：

```foxpro
*** 创建或获取wwdotnetbridge的缓存实例
LOCAL loBridge as wwDotnetBridge, loHttp
loBridge = GetwwDotnetBridge()

*** 创建一个内置的.NET类并运行一个方法
loHttp = loBridge.CreateInstance("System.Net.WebClient")
loHttp.DownloadFile("http://west-wind.com/files/MarkdownMonsterSetup.exe",
                    "MarkdownMonsterSetup.exe")
DO wwUtils
GoUrl(FULLPATH("MarkdownMonsterSetup.exe"))  && 运行它

*** 加载一个自定义 .NET 程序集
loBridge.LoadAssembly("CustomDotnet.dll")

*** 从新程序集中访问.NET组件
loItem = loBridge.CreateInstance("Custom.Item")

*** 直接访问属性
? loItem.Sku
loItem.Sku = "NewSku"
lnTotal = loItem.CalculateTotal()

*** 间接访问不可访问的属性和方法
lnFlagValue = loBridge.GetProperty(loItem,"Flag")
lnFlagValue = loBridge.SetProperty(loItem,"Flag",5) 
loBridge.InvokeMethod(loItem,"PassFlagValue",lnFlagValue)
```

## 功能一览
wwDotnetBridge 通过简单的 COM 互操作提供了以下增强功能：

* 直接访问大多数 .NET 组件，即使那些未标记[ComVisible]
* 不需要 COM 注册
* 调用任何 .NET 方法并设置和获取属性
* 以异步方式调用任何 .NET 方法，并在完成时回调
* 使用参数化构造函数实例化类型
* 使用标准方法名称调用重载方法
* 支持许多原生不支持的 .NET 类型和值
* 访问静态成员，值/结构类型，泛型，二进制，Guids，DbNulls
* 在方法返回时自动修复有问题的 .NET 类型
* 使用 ComArray 辅助类提供简单的数组访问
* ComValue 类在 .NET 中存储结果和参数
* ComValue 可帮助解决 .NET 类型转换问题
* ComArray 可以轻松创建，更新和管理 Enumerable 类型
* 内置多线程库
* wwDotnetBridge 还可以使用常规的 COM 互操作（不带运行时托管）

## 文档
* [主页](http://west-wind.com/wwDotnetBridge.aspx)
* [API 文档](https://www.west-wind.com/webconnection/wwClient_docs/_24n1cfw3a.htm)
* [白皮书](https://github.com/vfp9/wwDotnetBridge/blob/master/%E5%9C%A8%20Visual%20FoxPro%20%E4%B8%AD%E4%BD%BF%E7%94%A8%20%20wwDotnetBridge%20%20%E8%B0%83%E7%94%A8%20NET%20%E7%BB%84%E4%BB%B6.pdf)
* [更新历史](https://github.com/vfp9/wwDotnetBridge/blob/master/Changelog.md)

## 它是如何工作的
该库由3个组件组成（均在源代码中提供）：

* ClrHost.dll - .NET Runtime 的 Win32 载入器
* wwDotnetBridge.dll  - .NET 程序集的代理和增强
* wwDotnetBridge.prg - FoxPro 前端到 .NET 代理

> #### 确保可以找到DLL！
> 确保CrlHost.dll（或 wwIpstuff.dll 用于[商业West Wind工具](https://west-wind.com/WestwindClientTools.aspx)）和 wwDotnetBridge 可通过FoxPro路径访问。 理想情况下，您需要将这些DLL放在应用程序的当前执行路径中 - 通常是应用程序的根文件夹。

如果你使用的是FoxPro，你可以简单地使用wwDotnetBridge.prg类来加载一个程序集并且直接启动它或者访问现有的.NET组件。

此示例加载 OpenPop .NET 库并访问 POP3 邮箱以显示等待的消息：

```foxpro
*** 载入并初始化 wwDotnetBridge
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge")

*** 从磁盘载入一个程序集
loBridge.LoadAssembly("bin\OpenPop.dll")

*** 创建类实例 - 注意：没有 COM 注册
loPop = loBridge.CreateInstance("OpenPop.Pop3.Pop3Client")

*** 由于重载, 这不会起作用
* loPop.Connect("mail.server.net",587,.f.)

*** 因此，间接调用
? loBridge.InvokeMethod(loPop,"Connect","mail.server.net",110,.f.)

*** 大多数 方法/成员 都不会直接工作
? loPop.Authenticate("jb007","seekrit")

lnCount =  loPop.GetMessageCount()
? StringFormat("{0} Messages",lnCount)

*** 注意：OpenPop 从 1 开始，因为 pop3 是从 1 开始！
** 显示最后一条消息
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
