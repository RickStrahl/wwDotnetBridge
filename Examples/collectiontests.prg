CLEAR
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = GetwwDotnetBridge()

loNet = loBridge.Createinstance("Westwind.WebConnection.TypePassingTests")



*** Work with a List<T> 
? "*** Retrieve a List<T> and Update"

*** Generic List - can't do anything with this
*!*	loList = loNet.GetGenericList()
*!*	? loList && (Object)
*!*	? loList.Count && Nope

*** Indirect access through Proxy
loList = loBridge.InvokeMethod(loNet,"GetGenericList")
? loBridge.ToString(loList)   && System.Collections.Generic.List...
? loList.Count

*** Grab an item by index
loCust =  loList.Item(0)
? loCust.Company

*** Update item (updates underlying collection)
loCust.Company = loCust.Company + " " + TIME()

*** Create new item and add to collection
loCust = loList.CreateItem()
? loCust
loCust.Company = "North Wind Traders"
loList.AddItem(loCust)
? lolist.Count

*** Pass the ComArray instance back to .NET 
? loBridge.InvokeMethod(loNet, "SetGenericList", loList)  && should be 3 


*** Work with Dictionary<string, TestCustomer>
? "*** Retrieve a dictionary and Update"

loList = loBridge.InvokeMethod(loNet,"GetDictionary")
? loBridge.ToString(loList)
? loList.Count  &&  2

loCust =  loList.Item("Item1")   && Retrieve item by Key
? loCust.Company
loCust.Company = loCust.Company + " " + TIME()

loCust = loList.CreateItem()
? loCust
loCust.Company = "North Wind Traders"
loList.AddDictionaryItem("Item3", loCust)
? lolist.Count

? loBridge.InvokeMethod(loNet, "SetDictionary", loList)  && should be 3 

*** Get the raw Generic list - can't access this
*!*	? "-- Enumerate dictionary:"
*!*	loDictionary = loNet.GetDictionary()
*!*	? loDictionary
*!*	loComArray = loBridge.CreateInstance("Westwind.WebConnection.ComArray")
*!*	? loComArray.FromEnumerable(loDictionary)

*!*	? loComArray.Count

*!*	loItem =  loComArray.Item(1)
*!*	? loItem


RETURN
