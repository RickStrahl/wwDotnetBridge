
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = GetwwDotnetBridge()

? loBridge.GetDotnetVersion()

LOCAL loNet as Westwind.WebConnection.TypePassingTests
loNet = loBridge.Createinstance("Westwind.WebConnection.TypePassingTests")


*** Return a .NET STRUCT
LOCAL loComValue as "Westwind.WebConnection.ComValue"
loComValue = loBridge.CreateComValue()
loComArray = loBridge.CreateArray("System.Object")
loComArray.AddItem(1)
loComArray.AddItem("Hello World")
? loComValue.Value
loComValue.SetValueFromInvokeMethod(loNet,"ReturnStruct",loComArray)

*** Can't do this because structure is not accessible in VFP
*? loComValue.Value.IntValue
? loBridge.GetPropertyEx(loComValue,"Value.IntValue")
? loBridge.GetPropertyEx(loComValue,"Value.StringValue")

?
?

*** Return a .NET STRUCT
loComValue = loBridge.CreateComValue()
loComArray = loBridge.CreateArray("System.Object")
loComArray.AddItem(1)
loComArray.AddItem("Hello World")
? loComValue.Value
loComValue.SetValueFromInvokeStaticMethod("Westwind.WebConnection.TypePassingTests","ReturnStructStatic",loComArray)

*** Can't do this because structure is not accessible in VFP
*? loComValue.Value.IntValue
? loBridge.GetPropertyEx(loComValue,"Value.IntValue")
? loBridge.GetPropertyEx(loComValue,"Value.StringValue")


RETURN


*loNet.DecimalValue = 1.0
loBridge.SetProperty(loNet,"DecimalValue",1)
?loBridge.GetProperty(loNet,"DecimalValue")



*** PassLongArray
loLongs = loBridge.CreateArray("System.Int64")

loValue = loBridge.CreateComValue()
loValue.SetLong(2)
loLongs.AddItem(loValue)

loValue = loBridge.CreateComValue()
loValue.SetLong(5)
loLongs.AddItem(loValue)

loArray = loBridge.InvokeMethod(loNet,"PassLongArray",loLongs)
? loBridge.cErrorMsg
? loArray
? loArray.Count



*** Passing Single - Single is supported in VFP and direct calls work
***                  but you can't cast it in InvokeMethod calls
***                  so ComValue is required

*** Works
? loNet.PassSingle(10.2)

*** This doesn't work as FoxPro can't cast to Single
*? loBridge.InvokeMethod(loNet,"PassSingle",10.2)

*** PassSingle
loComValue = loBridge.CreateComValue()
loComValue.SetSingle(10.2)
? loBridge.InvokeMethod(loNet,"PassSingle",loComValue)



*** Passing Long - Longs are not supported in VFP natively
***                so conversion is ALWAYS required

*** Doesn't work
*? loNet.PassLong(10)
*? loBridge.InvokeMethod(loNet,"PassLong",10)

loComValue = lobridge.CreateComValue()
loComValue.SetLong(10)
? loBridge.InvokeMethod(loNet,"PassLong",loComValue)


*** Char by string
loComValue = loBridge.CreateComValue()
loComValue.SetChar("R")
? loBridge.InvokeMethod(loNet,"PassChar",loComValue)

*** Char by number
loComValue = loBridge.CreateComValue()
loComValue.SetChar(51)
? loBridge.InvokeMethod(loNet,"PassChar",loComValue)



RETURN

