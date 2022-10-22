CLEAR
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = GetwwDotnetBridge()

? loBridge.GetDotnetVersion()


loNet = loBridge.Createinstance("Westwind.WebConnection.TypePassingTests")
loList = loNet.GetGenericList()


*** Work with a List<T> 
? "*** Retrieve a List<T> and Update"
loList = loBridge.InvokeMethod(loNet,"GetGenericList")
? loBridge.ToString(loList)
? loList.Count

loCust =  loList.Item(0)
? loCust.Company
loCust.Company = loCust.Company + " " + TIME()

loCust = loList.CreateItem()
? loCust
loCust.Company = "North Wind Traders"
loList.AddItem(loCust)
? lolist.Count

? loBridge.InvokeMethod(loNet, "SetGenericList", loList)  && should be 3 


*** Work with Dictionary<string, TestCustomer>
? "*** Retrieve a dictionary and Update"

loList = loBridge.InvokeMethod(loNet,"GetDictionary")
? loBridge.ToString(loList)
? loList.Count  &&

loCust =  loList.Item("Item1")
? loCust.Company
loCust.Company = loCust.Company + " " + TIME()

loCust = loList.CreateItem()
? loCust
loCust.Company = "North Wind Traders"
loList.AddDictionaryItem("Item3", loCust)
? lolist.Count

? loBridge.InvokeMethod(loNet, "SetDictionary", loList)  && should be 3 

RETURN


loValue = loBridge.CreateComValue()

loValue.SetValueFromSystemConvert("ToInt32",22)
? loValue.ToString()  && 20  
? loValue.GetTypeName()
? loValue.Value

loValue.SetValueFromSystemConvert("ToUInt64","23")
? loValue.ToString()  && 20
? loValue.GetTypeName()  && System.UInt64

*** Long - Max Length
LOCAL loNet as Westwind.WebConnection.TypePassingTests
loNet = loBridge.Createinstance("Westwind.WebConnection.TypePassingTests")

? loBridge.InvokeMethod(loNet,"ReturnLong")



LOCAL loNet as Westwind.WebConnection.TypePassingTests
loNet = loBridge.Createinstance("Westwind.WebConnection.TypePassingTests")


loStrings = loBridge.CreateArray("System.String")
loStrings.AddItem("It's")
loStrings.AddItem("BigDay")

loValue = loBridge.CreateComValue(loStrings)

lobridge.InvokeMethod(loNet,"PassArrayByReference",loValue)

loStringsResult =  loValue.GetValue()
*loBridge.GetProperty(loValue,"Value")
? loStringsResult.Count


*** UInt64
LOCAL loU64 as Westwind.WebConnection.ComValue
loU64 = loBridge.CreateComValue(0)
loU64.SetUInt64(10)
? loBridge.SetProperty(loNet,"UInt64Value",loU64)
? loU64.ToString()

*** Long - Max Length
? "Long: " + TRANSFORM(loBridge.InvokeMethod(loNet,"ReturnLong"))


LOCAL loParms as WESTWIND.WebConnection.ComArray
loParms = loBridge.CreateArray("System.Object")
loParms.AddItem(INT(20))

loU64 = loBridge.CreateComValue()
loU64.SetValueFromInvokeStaticMethod("System.Convert","ToUInt64",loParms)
? loU64.ToString()



loInt = loBridge.CreateComValue(INT(10))
loString = loBridge.CreateComValue("Hello World.")
loDecimal = loBridge.CreateComValue(CAST( 5.22 as Y))
lobridge.InvokeMethod(loNet,"PassByReference",loInt,loString,loDecimal)

? loInt.Value, loString.Value, loDecimal.Value

loInt = loBridge.CreateComValue(INT(10))
loInt.Value = INT(10)
loString = loBridge.CreateComValue("Hello World.")
loDecimal = loBridge.CreateComValue(CAST( 5.22 as Currency))
lobridge.InvokeStaticMethod("Westwind.WebConnection.TypePassingTests",;
                            "PassByReferenceStatic",;
                            loInt,loString,loDecimal)

? loInt.Value, loString.Value, loDecimal.Value


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
? loBridge.GetProperty(loComValue,"Value.IntValue")
? loBridge.GetProperty(loComValue,"Value.StringValue")
? loBridge.InvokeMethod(loComValue,"Value.ToString")

?
?

*** Return a .NET STRUCT from a static method
loComValue = loBridge.CreateComValue()
loComArray = loBridge.CreateArray("System.Object")
loComArray.AddItem(1)
loComArray.AddItem("Hello World")
? loComValue.Value
loComValue.SetValueFromInvokeStaticMethod("Westwind.WebConnection.TypePassingTests","ReturnStructStatic",loComArray)

*** Can't do this because structure is not accessible in VFP
*? loComValue.Value.IntValue
? loBridge.GetProperty(loComValue,"Value.IntValue")
? loBridge.GetProperty(loComValue,"Value.StringValue")
? loBridge.InvokeMethod(loComValue,"Value.ToString")



*loNet.DecimalValue = 1.0
loBridge.SetProperty(loNet,"DecimalValue",CAST(1 as currency))
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


*** GUIDS cannot be accessed IN ANY WAY in FoxPro natively
*** wwDotnetBridge needs to wrpa them in a ComValue object
*** To retrieve a GUID always use `InvokeMethod` or `GetProperty`
loGuidComValue =  loBridge.InvokeMethod(loNet,"GetGuid") && returns COM Value
lcGuid = loGuidComValue.GetGuid() 
? lcGuid

lcGuid = loGuidComValue.NewGuid()
? "New Guid: " + lcGuid


*** Char by string
loComValue = loBridge.CreateComValue()
loComValue.SetChar("R")
? loBridge.InvokeMethod(loNet,"PassChar",loComValue)

*** Char by number
loComValue = loBridge.CreateComValue()
loComValue.SetChar(51)
? loBridge.InvokeMethod(loNet,"PassChar",loComValue)

*** Create Array On Instance
loCust = loBridge.CreateInstance("Westwind.WebConnection.TestCustomer")
? loCust             && Make sure you have what you think (Object)
? loCust.ToString()  && Make sure you have what you think 
? loBridge.ToString(loCust)  && for special types (structs, generics etc.) you have to use loBridge.ToString(loCust)

*** Clear out items that exist
loBridge.SetProperty(loCust,"Items",null)
? loCust.Items

? loBridge.Createarrayoninstance(loCust,"Items",2)
? loBridge.cErrorMsg


loItems = loBridge.GetProperty(loCust,"Items")
? loItems.Count


*** Work with a List<T> 
? "*** Retrieve a List<T> and Update"
loList = loBridge.InvokeMethod(loNet,"GetGenericList")
? loBridge.ToString(loList)
? loList.Count

loCust =  loList.Item(0)
? loCust.Company
loCust.Company = loCust.Company + " " + TIME()

loCust = loList.CreateItem()
? loCust
loCust.Company = "North Wind Traders"
loList.AddItem(loCust)
? lolist.Count

? loBridge.InvokeMethod(loNet, "SetGenericList", loList)  && should be 3 


*** Work with Dictionary<string, TestCustomer>
? "*** Retrieve a dictionary and Update"

loList = loBridge.InvokeMethod(loNet,"GetDictionary")
? loBridge.ToString(loList)
? loList.Count  &&

loCust =  loList.Item("Item1")
? loCust.Company
loCust.Company = loCust.Company + " " + TIME()

loCust = loList.CreateItem()
? loCust
loCust.Company = "North Wind Traders"
loList.AddDictionaryItem("Item3", loCust)
? lolist.Count

loItem =  loList.Item(0)
? loItem




RETURN
