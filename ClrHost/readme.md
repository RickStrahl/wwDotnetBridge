# ClrHost

This project contains the Win32 Stub loader for hosting the .NET Runtime in a Win23 process, specifically for use with wwDotnetBridge for hosting in Visual FoxPro.

There are two Win32 entry points:

* ClrCreateInstanceFrom() 
  Loads an assembly from its file path and then
  hosts the .NET runtime and passes back a pointer
  to a type of your choice. This type should be able
  to load other classes.

* ClrCreateInstance()
  Same as above but loads a strongly named, GAC'd
  assembly.


### How it works
These functions are Win32 loaders for a .NET type and a bootstrapper for .NET framework hosting in a Win32 process. The functions instantiate and return a .NET Type as a COM pointer, instantiating the .NET Rutime if it is not already running in the current process.

wwDotnetBridge uses this function to get a reference to a root wwDotnetBridge .NET object instance. This object can then be used to create instances of other .NET objects and interact with the .NET runtime system.

These functions are expensive to call so you should cache the retrieved reference and reuse it. wwDotnetBridge caches the object as an object property and globally if you use `GetwwDotnetBridge()` to retrieve the wwDotnetBridge reference.

### Example Code
The following example is the approximate loader code used by wwDotnetBridge, to instantiate the .NET wwDotnetBridge base object which is returned and turned into a COM reference:

```foxpro
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
```

If the .NET runtime is not loaded when this method is called these functions will load it.