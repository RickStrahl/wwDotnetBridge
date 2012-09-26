*** Resizes an image to a specified width 
*** using System.Drawing  GDI+ classes
*** 
***  Demonstrates:
***  =============
***  * Loads an assembly from the GAC
***  * Invoke Static Methods
***  * Access Static Properties
***  * Explicitly Invoke methods
CLEAR

*** Only for demo
SET PROCEDURE TO wwUtils ADDITIVE

do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge")

*** Load a .NET assembly from the GAC - use Reflector 
*** to inspect the full name on the DLL
loBridge.LoadAssembly("System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")

lcFileName = FULLPATH("sailbig.jpg")

loBmp = loBridge.Createinstance("System.Drawing.Bitmap",lcFileName)

*** Recalculate the size ratio
lnWidth = 150
lnHeight = (lnWidth/loBmp.Width) * loBmp.Height

*** Note: you have to cast numbers to the exact type
loBmp2 = loBridge.Createinstance("System.Drawing.Bitmap",;
                                 CAST(lnWidth as Integer),;
                                 CAST(lnHeight as Integer))


loGraphics = loBridge.Invokestaticmethod("System.Drawing.Graphics","FromImage",loBmp2)
? loBridge.cErRORMSG

loBrush = loBridge.GetStaticproperty("System.Drawing.Brushes","Black")

loBridge.InvokeMethod(loGraphics,"FillRectangle",;
						  loBrush,CAST(0 as Integer),CAST(0 as Integer),;
                          CAST(lnWidth as Integer),CAST(lnHeight as Integer))

loBridge.InvokeMethod(loGraphics,"DrawImage",;
                      loBmp,CAST(0 as Integer),CAST(0 as Integer),;
                      CAST(lnWidth as Integer),CAST(lnHeight as integer))

loGraphics.Dispose()

*** Make sure you release resources
loBmp.Dispose()

loBmp2.Save(FULLPATH("sailbig_resized.jpg"))
loBmp2.Dispose()

*** Show the image
GoUrl( FULLPATH("sailbig_resized.jpg"))
