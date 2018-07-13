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
该示例演示了一些简单的功能: 加载程序集, 创建. NET 类型实例, 然后直接或间接调用方法和访问属性。对于. NET 对象引用的许多方法和属性, 您可以直接访问成员, 但如果方法有重载, 如果方法或成员访问中涉及了泛型、枚举或值类型, 则某些成员不能直接通过 COM 调用。最好始终先尝试直接访问, 如果失败, 尝试使用对 wwDotnetBridge 实例的间接访问。

在这种情况下, 您可以使用间接引用来访问成员:

* loBridge.InvokeMethod(instance,"Method",parm1,parm2..,parm15)
* loBridge.GetProperty(instance,"Property")  
* loBridge.SetProperty(instance,"Property",valueToSet)
* GetPropertyEx(),SetPropertyEx,InvokeMethodEx() 支持成员的嵌套名称

这些方法在内部使用Reflection来调用.NET代码，但是因为它们在.NET内部运行，所以它们可以执行本机COM互操作所不能做的许多事情，因为COM上的类型编组的限制以及FoxPro类型系统的不兼容性。

这是另一个演示更多功能的示例，包括调用静态方法和设置枚举值以及使用枚举器访问集合。

```foxpro
*** 载入库
DO wwDotNetBridge

*** 创建 wwDotnetBridge 的实例
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")

loStore = loBridge.CreateInstance("System.Security.Cryptography.X509Certificates.X509Store")

? loBridge.cErrorMsg

*** 获取静态枚举值
leReadOnly = loBridge.GetEnumvalue("System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly")

*** 使用枚举值
loStore.Open(leReadOnly)   && 0 - 如果值是已知的

*** 返回存储项的. NET 集合
laCertificates = loStore.Certificates

*** 集合不适用于常规COM Interop
*** 所以使用间接访问
lnCount = loBridge.GetProperty(laCertificates,"Count")

*** 遍历 Certificates
FOR lnX = 0 TO lnCount -1
	*** 使用扩展语法间接访问集合项
	*** 支持嵌套对象和数组/集合 [] 括号
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

还可以异步调用任何. NET 方法, 并在方法完成执行后得到回调:

```foxpro
DO wwDotnetBridge

LOCAL loBridge as wwDotnetBridge
loBridge = GetwwDotnetBridge("V4")

loHttp = loBridge.CreateInstance("System.Net.WebClient")


*** 创建回调对象-对象必须是 "全局"
*** 所以当回调返回时, 它仍然存在
*** 使用Public或附加到长期存在的对象，如表单或_Screen
PUBLIC loCallback
loCallBack = CREATEOBJECT("HttpCallback")


*** 进行异步调用 - 立即返回
loBridge.InvokeMethodAsync(loCallback, loHttp, ;
                        "DownloadData",;
                        "http://west-wind.com/files/HelpBuilderSetup.exe")

? "下载已启动... 在后台运行。"

RETURN


*** 回调类 
DEFINE CLASS HttpCallback as Custom


FUNCTION OnCompleted(lvResult, lcMethod)

? "已完成 Http 调用"
? "收到： " + TRANS(LEN(lvResult),"999,999,999")

lcFile = ADDBS(SYS(2023)) + "HelpBuilderSetup.exe"

*** 写入到临时文件夹的文件
STRTOFILE(lvResult,lcFile)

*** 启动下载的安装程序
TRY
	*** 打开压缩文件
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

## 项目赞助商
以下人员/组织通过直接捐赠或付费开发为此项目提供赞助，作为使用这些工具的开发项目的一部分：

### West Wind Technologies
wwDotnetBridge最初是为[West Wind Client Tools](http://west-wind.com/webconnection)和[West Wind Web Connection](http://west-wind.com/wconnect)开发的，它们继续包括 略微修改了wwDotnetBridge版本。 West Wind Technologies通过允许与.NET轻松集成并允许更多人访问这一有用功能，友好地开源wwDotnetBridge以扩展FoxPro的范围。

wwDotnetBridge更新最初是为两种商业产品开发的，在进行更改时，任何更改都会合并到此项目中。 商业版本还包括产品使用的一些附加功能，例如SMTP客户端，SFTP支持，加密和图像管理实用程序，通过.NET包装器。 如果您想要一个完全支持的wwDotnetBridge版本，或者想赞助wwDotnetBridge的进一步开发工作，您可以通过购买这些产品的许可证来证明您的支持。

* [West Wind Web Connection](http://west-wind.com/webconnection)
* [West Wind Internet and Client Tools](http://west-wind.com/WestwindClientTools.aspx)


### Craig Tucker - Alabama Software
Craig offered early support and feedback for this project and billed project time for a number of additions to the library as part of a larger project.

### Bill Suthman - Monosynth
Bill为该项目提供了相当大的捐赠，并为一系列改进和错误修复提供了宝贵的反馈。

### Sunil Rjamara  - WeatherTrend
Sunil需要在他们的FoxPro产品中进行一些自定义集成，这导致发现了许多边缘案例，最终被集成到wwDotnetBridge中。 WeatherTrend捐赠了大量的工作时间来添加这些小功能。

### 想成为赞助商吗？
想赞助这个项目，需要定制还是想捐款来展示你的支持？ 您可以直接通过strahl@west-wind.com与我联系，或者您也可以通过PayPal在线捐款。

* [使用PayPal为wwDotnetBridge捐款](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=3CY6HGRTHSV5Y)
* [使用我们的网上商店为wwDotnetBridge捐款](http://store.west-wind.com/product/donation)
* [购买West Wind Internet和Client Tools的许可证](http://store.west-wind.com/product/wwclient50/)


## 许可证
该库根据**MIT许可**条款获得许可：

特此授予任何人获得本软件和相关文档文件 ( "软件 ") 副本的权限, 不受限制地处理软件, 包括不受限制地使用、复制、修改、合并的权利,发布、分发、再许可和/或销售软件的副本, 并允许软件所提供的人员这样做, 但须符合以下条件:

<small>&copy; 2012-2018 Rick Strahl, West Wind Technologies</small>

上述版权声明和本许可声明应包含在本软件的所有副本或实质部分中。

### 免责声明
该软件是 "按原样" 提供的, 没有任何明示或隐含的保证, 包括但不限于适销性、特定用途的适用性和适用性的保证。在任何情况下, 作者或版权持有人不得对任何索赔、损害赔偿或其他责任承担责任, 无论是在合同、侵权行为或其他方面, 因软件或软件的使用或其他交易而产生的或与之相关的。

<small>&copy; 2012-2018 Rick Strahl, West Wind Technologies</small>
