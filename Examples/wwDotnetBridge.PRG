SET PROCEDURE TO wwDotnetBridge ADDITIVE

*** This flag defines West Wind's commercial version
*** The open source version should be set to .F.
#DEFINE IS_WESTWIND        .F.
#DEFINE IS_LEGACY 	   .T.

#IF IS_WESTWIND
	#DEFINE WWC_CLR_HOSTDLL wwipstuff.dll 

	*** Other dependencies
	* .NET Runtime 2.0 or later

	*** DLLs
	* wwDotnetBridge.dll
	* wwIpStuff.dll
	* NewtonSoft.Json.dll - ToJson()/FromJson() only
#ELSE
	#DEFINE WWC_CLR_HOSTDLL clrhost.dll

	*** Other dependencies
	* .NET Runtime 2.0 or later

	*** DLLs
	* wwDotnetBridge.dll
	* clrhost.dll
#ENDIF


************************************************************************
*  GetwwDotNetBridge
****************************************
***  Function: Factory class to feed a reusable and stored instance
***            of the wwDotNetBridge class. 
***    Assume: Depends on wwutils.prg
***      Pass:
***    Return: 
************************************************************************
FUNCTION GetwwDotNetBridge(lcVersion)
IF VARTYPE(__DOTNETBRIDGE) != "O"
   *** Auto-detect latest version
   *** This requires wwUtils/wwApi
   IF EMPTY(lcVersion)       
	   lcVersion = GetHighestDotnetVersion()
	   IF ISNULL(lcVersion)
	       ERROR "Couldn't find a .NET version installed."
	   ENDIF
   ENDIF
   
   PUBLIC __DOTNETBRIDGE
   __DOTNETBRIDGE = CREATEOBJECT("wwDotNetBridge",lcVersion)
   IF VARTYPE(__DOTNETBRIDGE) # "O"
      RETURN NULL
   ENDIF
ENDIF
RETURN __DOTNETBRIDGE
ENDFUNC

************************************************************************
*  InitializeDotnetVersion()
****************************************
***  Function: Explicit method that simply creates a wwDotnetBridge
***            instance. Call this method at the beginning of your
***            app to force the .NET Runtime version that will be
***            used throughout the entire application.
***    Assume: Call this method in your application startup
***      Pass: "V4" or "V2"
***    Return: wwDotnetBridge instance 
***            (you don't have to do anything with it)
************************************************************************
FUNCTION InitializeDotnetVersion(lcVersion)
RETURN GetwwDotnetBridge(lcVersion)
ENDFUNC


*************************************************************
DEFINE CLASS wwDotNetBridge AS Custom
*************************************************************
*: Author: Rick Strahl
*:         (c) West Wind Technologies, 2007-2016
*:Contact: http://www.west-wind.com
*:Created: 05/10/2007
*:Updated: 6/06/2016
*************************************************************
#IF .F.
*:Help Documentation
*:Topic:
Class wwDotNetBridge

*:Description:
http://www.west-wind.com/webconnection/docs?page=_24n1cfw3a.htm

*:Example:

*:Remarks:

*:SeeAlso:


*:ENDHELP
#ENDIF

*** Custom Properties

*** Stock Properties
oDotNetBridge = null

cErrorMsg  = ""
lError = .f.

oLastException = null
FUNCTION oLastException_Access()
   RETURN this.oDotNetBridge.LastException
ENDFUNC

*** Default CLR Version. Note: you can override
*** this by specify V2,V4 in the Init() of the class
cClrVersion = "v4.0.30319"
lUseCom = .F.

*** Don't set here: Use CREATEOBJECT("wwDotNetBridge","V4")
*cClrVersion = "v4.0.30319"

************************************************************************
* wwDotnetBridge ::  Init
****************************************
***  Function: Initializes the .NET runtime inside of FoxPro
***    Assume: Calls Load() to actual perform load operation
***      Pass: lcVersion -  specific .NET Version or
***                         simplified "V2", "V4"
***            llUseCom  -  if .T. loads .NET Bridge via
***                         plain COM interop rather than 
***                         hosting .NET runtime itself.
***                         If you do COM 
***    Return: nothing
************************************************************************
FUNCTION Init(lcVersion, llUseCom)

IF !EMPTY(lcVersion)
   LOCAL lcShortVer
   lcShortVer = UPPER(lcVersion)
   DO CASE 
   	CASE lcShortVer == "V4"
   	  lcVersion = "v4.0.30319"
	CASE lcShortVer == "V2"
      lcVersion = "v2.0.50727"     
   ENDCASE

   this.cClrVersion = lcVersion
ENDIF   
this.lUseCom = llUseCom

*** Fail creation if the object couldn't be created
IF ISNULL(this.Load())
   ERROR "Unable to load wwDotNetBridge: " + this.cErrorMsg
   RETURN .F.
ENDIF

ENDFUNC
*  wwDotnetBridge ::  Init



************************************************************************
*  Load
****************************************
***  Function: Creates an instance of the .NET Bridge component
***            that allows loading of content.
***    Assume: Define as LOCAL obridge as Westwind.WebConnection.wwDotNetBridge
***      Pass:
***    Return:
************************************************************************
FUNCTION Load()

IF VARTYPE(this.oDotNetBridge) != "O"
	this.SetClrVersion(this.cClrVersion)

	IF this.lUseCom
		this.oDotNetBridge = CREATEOBJECT("Westwind.wwDotNetBridge")
	ELSE
		*** Load by filename - assumes the wwDotNetBridge.dll is in the Fox path
	   	DECLARE Integer ClrCreateInstanceFrom IN WWC_CLR_HOSTDLL string, string, string@, integer@

		lcError = SPACE(2048)
		lnSize = 0
		lnDispHandle = ClrCreateInstanceFrom(FULLPATH("wwDotNetBridge.dll"),;
				"Westwind.WebConnection.wwDotNetBridge",@lcError,@lnSize)

		IF lnDispHandle < 1
		   this.SetError( "Unable to load Clr Instance. " + LEFT(lcError,lnSize) )
		   RETURN NULL 
		ENDIF

		*** Turn handle into IDispatch COM object
		this.oDotNetBridge = SYS(3096, lnDispHandle)	

		*** Explicitly AddRef here - otherwise weird shit happens when objects are released
		SYS(3097, this.oDotNetBridge)

		IF ISNULL(this.oDotNetBridge)
			this.SetError("Can't access CLR COM reference.")
			RETURN null
		ENDIF
	ENDIF

	this.oDotNetBridge.LoadAssembly("System")	
ENDIF

RETURN this.oDotNetBridge
ENDFUNC
*   CreateDotNetBridge

************************************************************************
*  SetClrVersion
****************************************
***  Function: Sets the CLR Version that this runtime uses. Note
***            this method must be called before this.Load() is 
***            called and the runtime is instantiated for the very
***            first time.
***      Pass: lcVersion -  Clr version in the format of "v2.0.50727"
***    Return: nothing
************************************************************************
PROTECTED FUNCTION SetClrVersion(lcVersion)
DECLARE Integer SetClrVersion IN WWC_CLR_HOSTDLL string
SetClrVersion(lcVersion)
ENDFUNC
*   SetClrVersion

************************************************************************
*  Unload
****************************************
***  Function: Unloads the CLR and AppDomain
***    Assume: Don't call this unless you want to explicity force
***            the AppDomain to be unloaded and a new one to be spun
***            up. Generally a single domain shoudl be sufficient.
***      Pass:
***    Return:
************************************************************************
FUNCTION Unload()
IF VARTYPE(this.oDotNetBridge) == "O"
	this.oDotNetBridge = NULL	
	DECLARE Integer ClrUnload IN WWC_CLR_HOSTDLL	
	ClrUnload()
ENDIF
ENDFUNC
*   Unload

************************************************************************
*  CreateInstance
****************************************
***  Function: Creates an instance of a .NET class by full assembly
***            name or local file name.
***    Assume: Relies on wwDotNetBridgeW32.dll and wwDotNetBridge.dll
***      Pass: lcClass   - MyNamespace.MyClass
***            lvParmN   - Up to 3 Constructor Parameters
***    Return: Instance or NULL
************************************************************************
FUNCTION CreateInstance(lcClass,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5)
LOCAL lnDispHandle, lnSize, loBridge, loObject, lnParmCount

this.SetError()

loBridge = this.Load()
IF ISNULL(loBridge)
   RETURN NULL
ENDIF

lnParmCount = PCOUNT()
DO CASE 
	CASE lnParmCount = 2
		loObject = loBridge.CreateInstance_OneParm(lcClass,lvParm1)
    CASE lnParmCount = 3
        loObject = loBridge.CreateInstance_TwoParms(lcClass,lvParm1, lvParm2)
    CASE lnParmCount = 4
        loObject = loBridge.CreateInstance_ThreeParms(lcClass,lvParm1, lvParm2, lvParm3)
    CASE lnParmCount = 5
        loObject = loBridge.CreateInstance_FourParms(lcClass,lvParm1, lvParm2, lvParm3, lvParm4)
    CASE lnParmCount = 6
        loObject = loBridge.CreateInstance_FiveParms(lcClass,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5)        
    OTHERWISE
        loObject = loBridge.CreateInstance(lcClass)
ENDCASE

IF ISNULL(loObject) OR loBridge.Error
	this.SetError( loBridge.ErrorMessage )
	RETURN NULL
ENDIF

RETURN loObject
ENDFUNC

************************************************************************
*  CreateInstanceOnType
****************************************
***  Function: Creates an instance of a .NET class by full assembly
***            name or local file name.
***    Assume: Relies on wwDotNetBridgeW32.dll and wwDotNetBridge.dll
***      Pass: lcClass   - MyNamespace.MyClass
***            lvParmN   - Up to 5 Constructor Parameters
***    Return: Instance or NULL
************************************************************************
FUNCTION CreateInstanceOnType(loInstance, lcProperty, lcClass,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5)
LOCAL lnDispHandle, lnSize, loBridge, loObject, lnParmCount

this.SetError()

loBridge = this.Load()
IF ISNULL(loBridge)
   RETURN NULL
ENDIF

lnParmCount = PCOUNT()
DO CASE 
	CASE lnParmCount = 4
		loObject = loBridge.CreateInstanceOnType_OneParm(loInstance,lcProperty, lcClass,lvParm1)
    CASE lnParmCount = 5
        loObject = loBridge.CreateInstanceOnType_TwoParms(loInstance,lcProperty,lcClass,lvParm1, lvParm2)
    CASE lnParmCount = 6
        loObject = loBridge.CreateInstanceOnType_ThreeParms(loInstance,lcProperty,lcClass,lvParm1, lvParm2, lvParm3)
    CASE lnParmCount = 7
        loObject = loBridge.CreateInstanceOnType_FourParms(loInstance,lcProperty,lcClass,lvParm1, lvParm2, lvParm3, lvParm4)
    CASE lnParmCount = 8
        loObject = loBridge.CreateInstanceOnType_FiveParms(loInstance,lcProperty,lcClass,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5)
    OTHERWISE
        loObject = loBridge.CreateInstanceOnType(loInstance,lcProperty,lcClass)
ENDCASE

IF ISNULL(loObject) OR loBridge.Error
	this.SetError( loBridge.ErrorMessage )
	RETURN .F.
ENDIF

RETURN .T.
ENDFUNC


************************************************************************
*  InvokeMethod
****************************************
***  Function: Invokes an instance method on an object
***    Assume: Use when direct COM access doesn't work for things like collections/generics etc.
***      Pass:
***    Return:
************************************************************************
FUNCTION InvokeMethod(loObject, lcMethod, lvParm1, lvParm2, lvParm3, lvParm4, lvParm5,;
										  lvParm6, lvParm7, lvParm8, lvParm9, lvParm10,;
										  lvParm11, lvParm12, lvParm13, lvParm14, lvParm15,;
										  lvParm16, lvParm17, lvParm18, lvParm19, lvParm20,;
										  lvParm21, lvParm22, lvParm23, lvParm24)
LOCAL loBridge, lnParms, loResult

this.SetError()

loBridge = this.oDotNetBridge

lnParms = PCOUNT()
loResult = NULL
DO CASE
    CASE lnParms = 2
      loResult = loBridge.InvokeMethod(loObject, lcMethod)
	CASE lnParms = 3
	  loResult = loBridge.InvokeMethod_OneParm(loObject, lcMethod, lvParm1)
	CASE lnParms = 4
	  loResult = loBridge.InvokeMethod_TwoParms(loObject, lcMethod,lvParm1, lvParm2)
	CASE lnParms = 5
	  loResult = loBridge.InvokeMethod_ThreeParms(loObject, lcMethod,lvParm1, lvParm2, lvParm3)
	CASE lnParms = 6
	  loResult = loBridge.InvokeMethod_FourParms(loObject, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4)
	CASE lnParms = 7
	  loResult = loBridge.InvokeMethod_FiveParms(loObject, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5)
	CASE lnParms = 8
	  loResult = loBridge.InvokeMethod_SixParms(loObject, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6)
	CASE lnParms = 9
	  loResult = loBridge.InvokeMethod_SevenParms(loObject, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7)
	CASE lnParms = 10
	  loResult = loBridge.InvokeMethod_EightParms(loObject, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8)
	CASE lnParms = 11
	  loResult = loBridge.InvokeMethod_NineParms(loObject, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8, lvParm9)
	CASE lnParms = 12
	  loResult = loBridge.InvokeMethod_TenParms(loObject, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8, lvParm9, lvParm10)	  

	OTHERWISE
      	LOCAL loArray as Westwind.WebConnection.ComArray
	    loArray = this.CreateArray("System.Object")

	    LOCAL lvParm
	    FOR lnX = 1 TO lnParms-2
		   	lvParm = EVALUATE("lvParm" + TRANSFORM(lnX))
		   	loArray.AddItem(lvParm)   	
	    ENDFOR
	    
	    RETURN THIS.InvokeMethod_ParameterArray(loObject, lcMethod, loArray)
ENDCASE

IF loBridge.Error
   this.cErrorMsg = loBridge.ErrorMessage
   this.lError = .T.
ENDIF   

RETURN loResult
ENDFUNC
*   InvokeMethod


************************************************************************
*  InvokeMethod_ParameterArray
****************************************
***  Function: Invokes an instance method on an object
***    Assume: Use when direct COM access doesn't work for things like collections/generics etc.
***      Pass:
***    Return:
************************************************************************
FUNCTION InvokeMethod_ParameterArray
LPARAMETERS loObject, lcMethod, laParms
LOCAL loBridge, loResult, lnX

IF .F.
	LOCAL ARRAY laParms[1] && Compiler needs this or compile error
ENDIF

this.SetError()

LOCAL loBridge as Westwind.wwDotNetBridge
loBridge = this.oDotNetBridge

LOCAL loArray as Westwind.WebConnection.ComArray
loArray = this.CreateInstance("Westwind.WebConnection.ComArray")
loArray.Create("System.Object")

DO CASE
	CASE TYPE("laParms.Instance") # "U"
	   *** It's already a parameter array
	   loArray = laParms
	CASE TYPE("ALEN(laParms)") = "N"
	   *** It's a Fox array
	   lnsize = ALEN(laParms)
	   FOR lnX = 1 TO lnSize
		   	loArray.AddItem(laParms[lnX])   	
	   ENDFOR
	CASE TYPE("laParms.Count") = "N"
		*** It's a .NET array/collection
	    FOR lnX =1 TO laParms.Count
	    	loArray.AddItem(laParms[lnX])
	    ENDFOR
ENDCASE

loResult = loBridge.InvokeMethod_ParameterArray(loObject, lcMethod, loArray)

IF loBridge.Error
   this.cErrorMsg = loBridge.ErrorMessage
   this.lError = .T.
ENDIF   

RETURN loResult
ENDFUNC
*   InvokeMethod

************************************************************************
*  InvokeMethodAsync
****************************************
***  Function: Invokes a method in .NET Asynchronously on a new thread
***      Pass:  loCallbackEvents - CallbackAsyncEvents object
***                                Implement OnCompleted() and OnError()
***             loInstance       - Instance of object to call
***             lcMethod         - Method on instance to call
***             lvParm1-10       - 0 - 10 parameters
***    Return:  nothing
************************************************************************
FUNCTION InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod, lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8, lvParm9, lvParm10)
LOCAL lnParms

lnParms = PCOUNT()
LOCAL loBridge as wwDotNetBridge
loBridge = this.oDotnetBridge

DO CASE
  CASE lnParms = 3
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod)
  CASE lnParms = 4
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1)
  CASE lnParms = 5
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1,lvParm2)     
  CASE lnParms = 6
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1, lvParm2, lvParm3)     
  CASE lnParms = 7
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4)     
  CASE lnParms = 8
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5)          
  CASE lnParms = 9
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6)      
  CASE lnParms = 10
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7)      
  CASE lnParms = 11
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8)      
  CASE lnParms = 12
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8, lvParm9)      
  CASE lnParms = 13
     loBridge.InvokeMethodAsync(loCallbackEvents,loInstance,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8, lvParm9, lvParm10)           
ENDCASE

ENDFUNC
*   InvokeMethodAsync


************************************************************************
*  InvokeMethodAsync
****************************************
***  Function: Invokes a static method in .NET Asynchronously on a new thread
***      Pass:  loCallbackEvents - CallbackAsyncEvents object
***                                Implement OnCompleted() and OnError()
***             lcTypeName       - .NET Type on which static method lives
***             lcMethod         - Method on instance to call
***             lvParm1-10       - 0 - 10 parameters
***    Return:  nothing
************************************************************************
FUNCTION InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod, lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8, lvParm9, lvParm10)
LOCAL lnParms, loBridge as wwDotNetBridge
lnParms = PCOUNT()

loBridge = this.oDotnetBridge

DO CASE
  CASE lnParms = 3
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod)
  CASE lnParms = 4
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1)
  CASE lnParms = 5
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1,lvParm2)     
  CASE lnParms = 6
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1, lvParm2, lvParm3)     
  CASE lnParms = 7
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4)     
  CASE lnParms = 8
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5)          
  CASE lnParms = 9
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6)      
  CASE lnParms = 10
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7)      
  CASE lnParms = 11
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8)      
  CASE lnParms = 12
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8, lvParm9)      
  CASE lnParms = 13
     loBridge.InvokeStaticMethodAsync(loCallbackEvents,lcTypeName,lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5, lvParm6, lvParm7, lvParm8, lvParm9, lvParm10)           
ENDCASE

ENDFUNC
*   InvokeMethodAsync



************************************************************************
*  GetProperty
****************************************
***  Function: Gets a property from a .NET object
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetProperty(loInstance,lcProperty)
RETURN this.oDotNetBridge.GetProperty(loInstance, lcProperty) 
ENDFUNC
*   GetProperty

************************************************************************
*  GetIndexedProperty
****************************************
***  Function: Returns a value from an indexed property of an array
***            or a list.
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetIndexedProperty(loInstance, lnIndex)
RETURN this.oDotnetBridge.GetIndexedProperty(loInstance,lnIndex)
ENDFUNC
*   GetIndexedProperty

************************************************************************
*  SetProperty
****************************************
***  Function: Sets a property on a .NET object
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION SetProperty(loInstance, lcProperty, lvValue)
this.oDotNetBridge.SetProperty(loInstance, lcProperty, lvValue)
ENDFUNC
*   SetProperty


************************************************************************
*  InvokeStaticMethod
****************************************
***  Function: Calls a static .NET method with up to 5 parameters
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION InvokeStaticMethod(lcTypeName, lcMethod, lvParm1, lvParm2, lvParm3, lvParm4, lvParm5,;
                                                  lvParm6, lvParm7, lvParm8, lvParm9, lvParm10)
LOCAL loBridge, lnParms, loResult

this.SetError()

loBridge = this.oDotNetBridge

lnParms = PCOUNT()
loResult = NULL
DO CASE
	CASE lnParms = 3
	  loResult = loBridge.InvokeStaticMethod_OneParm(lcTypeName, lcMethod, lvParm1)
	CASE lnParms = 4
	  loResult = loBridge.InvokeStaticMethod_TwoParms(lcTypeName, lcMethod,lvParm1, lvParm2)
	CASE lnParms = 5
	  loResult = loBridge.InvokeStaticMethod_ThreeParms(lcTypeName, lcMethod,lvParm1, lvParm2, lvParm3)
	CASE lnParms = 6
	  loResult = loBridge.InvokeStaticMethod_FourParms(lcTypeName, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4)
	CASE lnParms = 7
	  loResult = loBridge.InvokeStaticMethod_FiveParms(lcTypeName, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5)
	CASE lnParms = 8
	  loResult = loBridge.InvokeStaticMethod_SixParms(lcTypeName, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5,lvParm6)
	CASE lnParms = 9
	  loResult = loBridge.InvokeStaticMethod_SevenParms(lcTypeName, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5,lvParm6, lvParm7)	
	CASE lnParms = 10
	  loResult = loBridge.InvokeStaticMethod_EightParms(lcTypeName, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5,lvParm6, lvParm7, lvParm8)	
	CASE lnParms = 11
	  loResult = loBridge.InvokeStaticMethod_NineParms(lcTypeName, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5,lvParm6, lvParm7, lvParm8, lvParm9)	
	CASE lnParms = 12
	  loResult = loBridge.InvokeStaticMethod_TenParms(lcTypeName, lcMethod,lvParm1, lvParm2, lvParm3, lvParm4, lvParm5,lvParm6, lvParm7, lvParm8, lvParm9, lvParm10)	

	OTHERWISE
	
	  loResult = loBridge.InvokeStaticMethod(lcTypeName, lcMethod)
ENDCASE

IF loBridge.Error
   this.cErrorMsg = loBridge.ErrorMessage
   this.lError = .T.
ENDIF   

RETURN loResult
ENDFUNC
*   InvokeStaticMethod

************************************************************************
*  GetStaticProperty
****************************************
***  Function: Retrieves a static property value from a type
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetStaticProperty(lcType, lcProperty)
LOCAL loBridge, lvResult

this.SetError()

loBridge = this.oDotNetBridge

lvResult = loBridge.GetStaticProperty(lcType, lcProperty)

IF loBridge.Error 
   this.SetError(loBridge.ErrorMessage)
   RETURN NULL
ENDIF   

RETURN lvResult
ENDFUNC
*   GetStaticProperty


************************************************************************
*  GetEnumValue
****************************************
***  Function: Retrieves an Enum value as a number from a string representation
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetEnumValue(lcType, lcProperty)

IF !EMPTY(lcProperty)
  RETURN this.GetStaticProperty(lcType, lcProperty)
ENDIF

*** Assume we're passing the type and enum name as one string
lnAt = RAT(".",lcType)
IF lnAt < 1
   RETURN NULL
ENDIF

lcProperty = SUBSTR(lcType,lnAt+1)
lcType = SUBSTR(lcType,1,lnAt-1)

RETURN this.GetStaticProperty(lcType, lcProperty)
ENDFUNC
*   GetEnumValue


************************************************************************
* wwDotnetBridge ::  GetEnumString
****************************************
***  Function: Retrieves the enum field name from an enum value
***    Assume:
***      Pass: lvEnumValue - The Enum value to retrieve string for
***    Return:
************************************************************************
FUNCTION GetEnumString(lcEnumTypeName,lvEnumValue)
LOCAL lcValue

lcValue = THIS.oDotNetBridge.GetEnumString(lcEnumTypeName,lvEnumValue)
IF ISNULL(lcValue)
   this.SetError(this.oDotNetBridge.ErrorMessage)
   RETURN NULL
ENDIF   

RETURN lcValue
ENDFUNC
*  wwDotnetBridge ::  GetEnumString

************************************************************************
*  SetStaticProperty
****************************************
***  Function: Sets a static property value on a .NET type
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION SetStaticProperty(lcType,lcProperty,lcValue)
LOCAL loBridge, lvResult

loBridge = this.oDotNetBridge

IF !loBridge.SetStaticProperty(lcType, lcProperty, lcValue) 
   this.SetError(loBridge.ErrorMessage)
   RETURN .F.
ENDIF   

RETURN .T.
ENDFUNC
*   SetStaticProperty



************************************************************************
*  CreateArray
****************************************
***  Function: Creates an array of a given type and size on the specified
***            object instance.
***    Assume:
***      Pass:
***    Return: ComArray object with Instance,Length properties or null
************************************************************************
FUNCTION CreateArray(lvArrayInstanceOrElementTypeString)
LOCAL loComArray, lcType

lcType = VARTYPE(lvArrayInstanceOrElementTypeString)
DO CASE
  CASE lcType = "X" OR EMPTY(lvArrayInstanceOrElementTypeString)
	  loComArray = this.oDotnetBridge.CreateInstance("Westwind.WebConnection.ComArray")
  CASE lcType = "C"
	  loComArray = this.oDotnetBridge.CreateArray(lvArrayInstanceOrElementTypeString)
  CASE lcType = "O"
	  loComArray = this.oDotnetBridge.CreateArrayFromInstance(lvArrayInstanceOrElementTypeString)
  OTHERWISE
      THROW "Invalid parameter passed to CreateArray"
ENDCASE

IF ISNULL(loComArray)
   THIS.SetError(this.oDotnetBridge.ErrorMessage)
   RETURN null
ENDIF

RETURN loComArray
ENDFUNC
*   CreateArray

************************************************************************
*  CreateComValue
****************************************
***  Function: Creates a COM Value structure
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION CreateComValue(lvValue)
LOCAL loVal
loVal = this.CreateInstance("Westwind.WebConnection.ComValue")

if(PCOUNT()=1)
   this.SetProperty(loVal,"Value",lvValue)
ENDIF
RETURN loVal
ENDFUNC
*   CreateComValue


************************************************************************
*  CreateArrayOnInstnace
****************************************
***  Function: Creates an array of a given type and size on the specified
***            object instance.
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION CreateArrayOnInstance(loBaseType, lcType, lnSize)

IF VARTYPE(lnSize) # "N"
	IF !this.oDotNetBridge.CreateArrayOnInstanceWithObject(loBaseType, lcType, lnSize)
	   THIS.SetError(this.oDotnetBridge.ErrorMessage)
	   RETURN .F.
	ENDIF
	RETURN .T.	
ENDIF

IF !this.oDotNetBridge.CreateArrayOnInstance(loBaseType, lcType, lnSize)
   THIS.SetError(this.oDotnetBridge.ErrorMessage)
   RETURN .F.
ENDIF

RETURN .T.
ENDFUNC
*   CreateArray

************************************************************************
*  AddArrayItem
****************************************
***  Function: Creates an item for an array and adds it to the array
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION AddArrayItem(loBaseType, lcArrayProperty, loValue)

IF !this.oDotNetBridge.AddArrayItem(loBaseType, lcArrayProperty,loValue)
   this.SetError("Couldn't add item to array: " + this.oDotNetBridge.ErrorMessage)
   RETURN .F.
ENDIF

RETURN .T.
ENDFUNC
*   AddArrayItem

************************************************************************
*  GetArrayItem
****************************************
***  Function: Returns an individual array item.
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetArrayItem(loBaseType, lcArrayProperty, lnIndex)
RETURN this.oDotNetBridge.GetArrayItem(loBaseType, lcArrayProperty,lnIndex)
ENDFUNC
*   GetArrayItem

************************************************************************
*  SetArrayItem
****************************************
***  Function: Sets an individual array item.
***    Assume: The array item must exist - ie. array is big enough
***      Pass:
***    Return:
************************************************************************
FUNCTION SetArrayItem(loBaseType, lcArrayProperty, lnIndex, lvValue)
RETURN this.oDotNetBridge.SetArrayItem(loBaseType, lcArrayProperty,lnIndex,lvValue)
ENDFUNC
*   SetArrayItem



************************************************************************
*  RemoveArrayItem
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION RemoveArrayItem(loBaseType, lcArrayProperty, lnIndex)

IF !this.oDotNetBridge.RemoveArrayItem(loBaseType,lcArrayProperty, lnIndex)
	this.SetError("Couldn't remove item from array: " + this.oDotnetBridge.ErrorMessage)
	RETURN .F.
ENDIF

RETURN .T.	
ENDFUNC
*   RemoveArrayItem


************************************************************************
*  GetType
****************************************
***  Function: Returns a .NET type reference to the value passed
***    Assume: Note this explicit method is required in order
***            to call GetType() 
***      Pass:
***    Return:
************************************************************************
FUNCTION GetType(lvValue)
RETURN this.InvokeMethod(this.oDotnetBridge,"GetType",lvValue)
ENDFUNC
*   GetType

************************************************************************
*  GetTypeFromName
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetTypeFromName(lcTypeName)
RETURN this.InvokeMethod(this.oDotNetBridge,"GetTypeFromName",lcTypeName)
ENDFUNC
*   GetTypeFromName


************************************************************************
*  DataSetToXmlAdapter
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION DataSetToXmlAdapter(loDs)
LOCAL lcXML 

lcXML = this.InvokeMethod(this.oDotNetBridge,"DataSetToXmlString",loDs,.T.)

IF EMPTY(lcXml)
	RETURN null
ENDIF

LOCAL loAdapter as XMLAdapter
loAdapter = CREATEOBJECT("XmlAdapter")
loAdapter.LoadXML(lcXml,.F.,.T.)

RETURN loAdapter
ENDFUNC
*   DataSetToXmlString

************************************************************************
*  DataSetToCursors
**************************************** 
***  Function: Converts a .NET DataSet to cursors that match the
***            tables in the Tables collection.
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION DataSetToCursors(loDs)

loAdapter = this.DataSetToXmlAdapter(loDs)
IF ISNULL(loAdapter)
   RETURN .F.
ENDIF   

RETURN this.XmlAdapterToCursors(loAdapter)
ENDFUNC
*   DataSetToCursors


************************************************************************
*  XmlStringToDataSet
****************************************
***  Function: Converts an XML String created with XmlAdapter or
***            CursorToXml into a DataSet
***    Assume:
***      Pass: lcXml  -  Xml string
***    Return: .NET DataSet object or NULL
************************************************************************
FUNCTION XmlStringToDataSet(lcXml)
LOCAL loDataSet

loDataSet = this.oDotNetBridge.XmlStringToDataSet(lcXml)

IF ISNULL(loDataSet)
   this.SetError(this.oDotNetBridge.ErrorMessage)
   RETURN null
ENDIF

RETURN loDataSet
ENDFUNC
*   XmlStringToDataSet

************************************************************************
*  CursorToDataSet
****************************************
***  Function: 
***    Assume:
***      Pass: lcAliasList - A list of comma delimited cursors to add to 
***                          the data
***    Return: DataSet or NULL
************************************************************************
FUNCTION CursorToDataSet(lcAliasList)
LOCAL lnX, lcXml

IF EMPTY(lcAliasList)
   RETURN null
ENDIF

LOCAL loAdapter as XMLAdapter
loAdapter = CREATEOBJECT("XmlAdapter")

LOCAL ARRAY laCursors[1]   
lnCount = ALINES(laCursors,lcAliasList,1 + 4,",")
FOR lnX = 1 TO lnCount
	loAdapter.AddTableSchema(laCursors[lnX],.T.)
ENDFOR

lcXml = ""
loAdapter.ToXML("lcXml")

IF EMPTY(lcXml)
   RETURN null
ENDIF

RETURN THIS.XmlStringToDataSet(lcXml)   
ENDFUNC
*   CursorToDataSet

************************************************************************
*  XmlAdapterToCursors
****************************************
***  Function: Opens all tables in an XmlAdapter as Cursors
***    Assume:
***      Pass:
***    Return: .T. or .F.
************************************************************************
FUNCTION XmlAdapterToCursors(loAdapter)

IF VARTYPE(loAdapter) # "O" 
   this.SetError("No Adapter passed")
   RETURN .F.	
ENDIF
IF loAdapter.Tables.Count < 1
	THIS.SetError("No tables on XmlAdapter")
	RETURN .F.
ENDIF
FOR EACH loTable IN loAdapter.Tables
    USE IN SELECT(loTable.Alias)
    loTable.ToCursor(.F.,loTable.Alias)
ENDFOR

RETURN .T.
ENDFUNC
*   XmlAdapterToTables

************************************************************************
*  XmlAdapterGetCursor
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION XmlAdapterGetCursor(loAdapter,lvCursor)
LOCAL loTable, loTbl

IF EMPTY(lvCursor)
   *** Get the first cursor if no cursor name is passed
   IF loAdapter.Tables.Count > 0
   	  lvCursor = loAdapter.Tables[1].Alias
   ELSE
   	  RETURN .F.
   ENDIF
ENDIF   

loTable = null

IF VARTYPE(lvCursor) = "N"
   loTable = loAdapter.tables[lvCursor]
ELSE
	lvCursor = LOWER(lvCursor)
	FOR EACH loTbl in loAdapter.Tables
	   IF lvCursor == LOWER(loTbl.alias)
	      loTable = loTbl
	      EXIT
	   ENDIF      
	ENDFOR
ENDIF

IF ISNULL(loTable)
   RETURN .F.
ENDIF   

USE IN SELECT(loTable.Alias)
  
loTable.ToCursor(.f.,loTable.Alias)
RETURN .T.
ENDFUNC
*   XmlAdapterGetCursor


************************************************************************
*  LoadAssembly
****************************************
***  Function: Loads an assembly either by full assembly name or by 
***            fully qualified path.
***    Assume:
***      Pass: AssemblyName or full assembly file name
***    Return: .T. or .F.
************************************************************************
FUNCTION LoadAssembly(lcAssembly)
LOCAL loBridge

this.SetError()

loBridge = this.Load()

IF ISNULL(loBridge)
   RETURN NULL
ENDIF

IF FILE(lcAssembly)
	lcAssembly = LOWER(FULLPATH(lcAssembly))
ENDIF

IF AT(":\",lcAssembly) > 0 OR AT("\\",lcAssembly) > 0
	*** It's a file based assembly path
	IF !loBridge.LoadAssemblyFrom(lcAssembly)
		this.SetError(lobridge.ErrorMessage)
		RETURN .F.
	ENDIF
ELSE
    *** It's an assembly name
	IF !loBridge.LoadAssembly(lcAssembly)
		this.SetError(lobridge.ErrorMessage)
		RETURN .F.
	ENDIF
ENDIF

RETURN .T.	
ENDFUNC
*   LoadAssembly

************************************************************************
* wwDotNetBridge ::  RunThread
****************************************
***  Function: Runs a PRG file on a separate thread
***    Assume:
***      Pass: loEvents - Any Fox object that can be called
***                       from within the PRG
***    Return: Instance of the ThreadRunner .NET object
************************************************************************
FUNCTION RunThread(lcPrgFileName, loEvents)
LOCAL loThread
loThread = THIS.CreateInstance("Westwind.WebConnection.ThreadRunner")
loThread.RunThread(loEvents,FULLPATH(lcPrgFileName))
RETURN loThread
ENDFUNC
*  wwDotNetBridge ::  ExecutePrgOnThread

*!*	************************************************************************
*!*	*  CreateClrInstanceFrom
*!*	****************************************
*!*	***  Function: Creates an instance of a .NET class by referencing
*!*	***            a fully qualified assembly path.
*!*	***      Pass: lcLibrary - fully qualified path to the assembly 
*!*	***                        including extension
*!*	***            lcClass   - MyNamespace.MyClass
*!*	***    Return: Instance or NULL
*!*	************************************************************************
*!*	FUNCTION CreateClrInstance(lcLibrary,lcClass,lcError)
*!*	LOCAL lnDispHandle, lnSize

*!*	DECLARE Integer ClrCreateInstance IN WWC_CLR_HOSTDLL string, string, string@, integer@

*!*	lcError = SPACE(2048)
*!*	lnSize = 0
*!*	lnDispHandle = ClrCreateInstance("wcDotNetBridge","Westwind.WebConnection.wwDotNetBridge",@lcError,@lnSize)

*!*	IF lnDispHandle < 1
*!*	   lcError = LEFT(lcError,lnSize)
*!*	   RETURN NULL 
*!*	ENDIF

*!*	RETURN SYS(3096, lnDispHandle)
*!*	ENDFUNC


************************************************************************
*  ConvertToDotnetValue
****************************************
***  Function: 
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ConvertToDotnetValue(lvFoxValue,lcType)

lcType = LOWER(lcType)

*** Only create loVal object on types that need conversion
IF EMPTY(lcType) OR INLIST(lcType,"int16","int64","byte","dbnull")
	LOCAL loVal as Westwind.WebConnnection.ComValue
	loVal = this.CreateInstance("Westwind.WebConnection.ComValue")
ENDIF

IF EMPTY(lcType)
   RETURN loVal
ENDIF

   
DO CASE
  *** ComValue object conversions
  CASE lcType = "int16"
      loVal.SetInt16(lvFoxValue)
  CASE lcType = "int64"
      loVal.SetInt64(lvFoxValue)
  CASE lcType = "byte"
      loVal.SetByte(lvFoxValue)
  CASE lcType = "char"
      loVal.SetChar(lvFoxValue)
  CASE lcType = "dbnull"
      loVal.SetDbNull()     

  *** CAST Conversions
  CASE lcType = "decimal"
      RETURN NTOM(lvFoxValue)
      *CAST(lvFoxValue as Currency)      
  CASE lcType = "binary" OR lcType = "byte[]"
      #IF VERSION(5) > 800
	     RETURN CAST(lvFoxValue as BLOB)      
	  #ELSE
	  	 RETURN CREATEBINARY(lvFoxValue)
      #ENDIF
  OTHERWISE
  	  *** If no conversion exists just return the value as is
      RETURN lvFoxValue
ENDCASE

RETURN loVal
ENDFUNC
*   ConvertToDotNetValue


************************************************************************
*  GetDotNetVersion
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetDotNetVersion()
RETURN this.oDotNetBridge.GetVersionInfo()
ENDFUNC
*   GetDotNetVersion

************************************************************************
*  SetError
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION SetError(lcMessage)

IF EMPTY(lcMessage)
   this.lError = .f.
   this.cErrorMsg = ""
   RETURN
ENDIF

this.lError = .t.
this.cErrorMsg = lcMessage   

ENDFUNC
*   SetError



#IF IS_WESTWIND

************************************************************************
*  ToJson
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ToJson(loDotnetObject, llFormatted)

lcResult = this.oDotnetBridge.ToJson(loDotnetObject,llFormatted)
IF ISNULL(lcResult)
   THIS.SetError( this.oDotnetBridge.ErrorMessage)
ENDIF

RETURN lcResult
ENDFUNC
*   ToJson


************************************************************************
*  FromJson
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION FromJson(lcJson,lvType)

loType = lvType
IF VARTYPE(lvType) = "C"
   loType = this.GetTypeFromName(lvType)
ENDIF

lcResult = this.oDotnetBridge.FromJson(lcJson,loType)   
IF ISNULL(lcResult)
   this.SetError(this.oDotnetBridge.ErrorMessage)
   RETURN NULL
ENDIF   

RETURN lcResult
ENDFUNC
*   FromJson

************************************************************************
*  ToXml
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ToXml(loDotnetObject, llFormatted)

lcResult = this.oDotnetBridge.ToXml(loDotnetObject)
IF ISNULL(lcResult)
   THIS.SetError( this.oDotnetBridge.ErrorMessage)
ENDIF

RETURN lcResult
ENDFUNC
*   ToXml

************************************************************************
*  FromXml
****************************************
***  Function:
***    Assume:
***      Pass: lcXml    -   Xml string
***            lvType   -  .NET type as string or .NET type object
***    Return:
************************************************************************
FUNCTION FromXml(lcXml,lvType)

loType = lvType
IF VARTYPE(lvType) = "C"
   loType = this.GetTypeFromName(lvType)
ENDIF

lcResult = this.oDotnetBridge.FromXml(lcXml,loType)   

IF ISNULL(lcResult)
   this.SetError(this.oDotnetBridge.ErrorMessage)
   RETURN NULL
ENDIF   

RETURN lcResult
ENDFUNC
*   FromXml

#ENDIF

#IF IS_LEGACY
************************************************************************
*  GetPropertyEx
****************************************
***  Function: Gets a property from a .NET object using .syntax for 
***            for the property as well as supporting arrays
***    Assume: OBSOLETE - just use 
***      Pass: 
***    Return:
************************************************************************
FUNCTION GetPropertyEx(loInstance,lcProperty)
RETURN this.oDotNetBridge.GetPropertyEx(loInstance, lcProperty)
ENDFUNC
*   GetPropertyEx

************************************************************************
*  SetPropertyEx
****************************************
***  Function: Sets a property on a .NET object 
***    Assume: OBSOLETE - Just use GetProperty()
***      Pass:
***    Return:
************************************************************************
FUNCTION SetPropertyEx(loInstance, lcProperty, lvValue)
this.oDotNetBridge.SetPropertyEx(loInstance, lcProperty, lvValue)
ENDFUNC
*   SetProperty
#ENDIF


ENDDEFINE
*EOC wwDotNetBridge 


*************************************************************
DEFINE CLASS AsyncCallbackEvents AS Custom
*************************************************************
*: Author: Rick Strahl
*:         (c) West Wind Technologies, 2015
*:Contact: http://www.west-wind.com
*:Created: 12/23/2015
*************************************************************
#IF .F.
*:Help Documentation
*:Topic:
Class AsyncCallbackEvents

*:Description:
Create a subclass of this class and pass to InvokeMethodAsync
as a callback object. When the method call completes or fails
the OnCompleted or OnError methods are fired.

Make sure this object stays in scope for the duration of
the async call - otherwise the callback will fail.

*:Example:

*:Remarks:

*:SeeAlso:


*:ENDHELP
#ENDIF


************************************************************************
*  OnCompleted
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION OnCompleted(lvResult,lcMethod)
ENDFUNC
*   OnCompleted

************************************************************************
*  OnError
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION OnError(lcMessage, loException,lcMethod)
ENDFUNC
*   OnError

ENDDEFINE

*EOC AsyncCallbackEvents 



************************************************************************
* GetHighestDotnetVersion
****************************************
***  Function: Returns the highest .NET version installed on the
***            local machine.
************************************************************************
FUNCTION GetHighestDotnetVersion()
LOCAL x,lcFrameworkPath, lcVersion

lcVersion = ""
lcFrameworkPath = SPACE(256)

DECLARE INTEGER GetWindowsDirectory ;
   IN Win32API ;
   STRING  @pszSysPath,;
   INTEGER cchSysPath
lnsize=GetWindowsDirectory(@lcFrameworkPath,256) 
if lnSize > 0
   lcFrameworkPath = SUBSTR(lcFrameworkPath,1,lnSize) + "\"
ELSE
   lcFrameworkPath = TRIM(lcFrameworkPath)
ENDIF
   
*** Assume .NET 2.0
lcVersion = NULL
   
   *** Try to find the largest version number
lcFrameworkPath = lcFrameworkPath + "Microsoft.NET\Framework\"
lnCount = ADIR(laNetDirs,lcFrameworkPath + "v?.*.*","D")
IF lnCount < 1
	RETURN null
ENDIF

*** Highest version comes last so go backwards through list
FOR x = lnCount TO 1 STEP -1
   lcVersion = laNetDirs[x,1]
   lcTPath = ADDBS(lcFrameworkPath + lcVersion )
   IF FILE(lcTPath + "regasm.exe")
     lcFrameworkPath = ADDBS(lcTPath)         
     EXIT
   ENDIF
ENDFOR   

RETURN lcVersion
ENDFUNC
*  GetHighestDotnetVersion
