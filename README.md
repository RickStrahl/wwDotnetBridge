#wwDotnetBridge
==============

wwDotnetBridge is a .NET Runtime Hosting engine and .NET access helper for Visual FoxPro. It allows hosting of the .NET runtime in a Win32 application, and accessing .NET components without having to have them registered first.

It provides the following enhancements over plain COM Interop:

* Access most .NET Components directly
* Requires no COM Registration for .NET Objects
* Create types with parameterized constructors
* Support for many natively unsupported .NET types and values
* Access static method/members, structs, binary data, guids, DbNulls
* Call overloaded methods 
* Automatically fixes up problematic .NET Types
* Provides easy Array access with ComArray helper class
* ComValue structure to help with .NET TypeConversions
* Ability to keep values entirely in .NET with ComValue
* Assign values directly to properties without passing through FoxPro
* Multi-threading library built-in

Online Documentation:
* [Documentation](http://www.west-wind.com/webconnection/wwClient_docs/_24n1cfw3a.htm)

##How it works

This library consists of 3 components:

* ClrHost.dll - Win32 Loader for the .NET Runtime
* wwDotnetBridge.dll  - .NET assembly Proxy and Helper
* wwDotnetBridge.prg - FoxPro front end to .NET Proxy

If you're using FoxPro you can simply use the wwDotnetBridge.prg class to load an assembly and fire away at it or access existing .NET components directly.

This example loads the OpenPop .NET library and accesses a POP3 mailbox to display messages waiting:

    *** Load library and initialize wwDotnetBridge
	do wwDotNetBridge
	LOCAL loBridge as wwDotNetBridge
	loBridge = CreateObject("wwDotNetBridge")

	*** Load an assembly from disk
	loBridge.LoadAssembly("bin\OpenPop.dll")

    *** Create an instance of a class - note: No COM registration
	loPop = loBridge.CreateInstance("OpenPop.Pop3.Pop3Client")

	*** This won't work due to overloads
	* loPop.Connect("mail.gorge.net",587,.f.)

	*** So, call indirectly instead
	? loBridge.InvokeMethod(loPop,"Connect","pop3.gorge.net",110,.f.)

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

The example demonstrates a few simple features: Loading an assembly, creating a .NET type instance and then calling methods and accessing properties either directly or indirectly. For many methods and properties on .NET object references you can directly access the members, but some members are not directly callable via COM if there are overloads on a method, if there are Generics, enums or Value Types involved in the method or member access. It's best to always try direct access first and if that fails attempt to use indirect access to the wwDotnetBridge instance.

In that case you can use indirect referencing to access members with:

* loBridge.InvokeMethod(instance,"Method",parm1,parm2..,parm15)
* loBridge.GetProperty(instance,"Property")  
* loBridge.SetProperty(instance,"Property",valueToSet)
* GetPropertyEx(),SetPropertyEx,InvokeMethodEx() which supported nested names for the member


These methods internally use Reflection to call .NET code, but because they run inside of .NET they can do thing that native COM interop cannot. 

Here's another example that demonstrates a few more features including calling static methods and setting enum values and accessing collections using enumerators.

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

There's much more functionality available. Please check out the documentation for more info.

* [Documentation](http://www.west-wind.com/webconnection/wwClient_docs/_24n1cfw3a.htm)
