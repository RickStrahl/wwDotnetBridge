#ClrLoader.dll
-------------
This project contains the Win32 Stub loader for hosting the .NET Runtime

There are two Win32 entry points:

* ClrCreateInstanceFrom() 
  Loads an assembly from its file path and then
  hosts the .NET runtime and passes back a pointer
  to a type of your choice. This type should be able
  to load other classes.

* ClrCreateInstance()
  Same as above but loads a strongly named, GAC'd
  assembly.

From FoxPro use something like this to load

	*** Load by filename - assumes the wwDotNetBridge.dll is in the Fox path
   	DECLARE Integer ClrCreateInstanceFrom IN ClrHost.dll ;
                           string, string, string@, integer@
	
	lcError = SPACE(2048)
	lnSize = 0
	lnDispHandle = ClrCreateInstanceFrom(FULLPATH("wwDotNetBridge.dll"),;
			"Westwind.WebConnection.wwDotNetBridge",;  && .NET Type to load
                        @lcError,@lnSize)
	
	IF lnDispHandle < 1
	   this.SetError( "Unable to load Clr Instance. " + LEFT(lcError,lnSize) )
	   RETURN NULL 
	ENDIF
	
	*** Turn handle into IDispatch object
	this.oDotNetBridge = SYS(3096, lnDispHandle)	
	
	*** Explicitly AddRef here - otherwise weird shit happens when objects are released
	SYS(3097, this.oDotNetBridge)
	
	IF ISNULL(this.oDotNetBridge)
		this.SetError("Can't access CLR COM reference.")
		RETURN null
	ENDIF
