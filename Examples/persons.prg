CLEAR
DO wwutils && Demo only
*LOCAL loFox as InteropExamples.Examples
*loFox = CREATEOBJECT("InteropExamples.Examples")


do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")

*** Load our custom assembly
loBridge.LoadAssembly("InteropExamples.dll")
IF loBridge.lError
   ? loBridge.cErrorMsg
ENDIF

loFox = loBridge.CreateInstance("InteropExamples.Examples")
IF loBridge.lError
   ? loBridge.cErrorMsg
ENDIF

loPerson = loFox.GetNewPerson()
loPerson.Id = 1
loPerson.FirstName = "Rick"
loPerson.LastName = "Strahl"
loPerson.Address.Street = "32 Kaiea Place"
loPerson.Entered = DATETIME()
loFox.SavePerson(loPerson)

loPerson = loFox.GetNewPerson()
loPerson.Id = 2
loPerson.FirstName = "Markus"
loPerson.LastName = "Egger"
loPerson.Address.Street = "213 Mud Lane"
loPerson.Entered = DATETIME()
loFox.SavePerson(loPerson)

loPerson = null

loPerson = loFox.GetPerson(1)
? loPerson.FirstName + " " + loPerson.Address.Street

*!*	? loFox.Persons  && (Object)
*!*	? loFox.Persons.Count  && Error

*!*	loPersons = loFox.GetPersons()
*!*	? loPersons[1].FirstName


*!*	do wwDotNetBridge
*!*	LOCAL loBridge as wwDotNetBridge
*!*	loBridge = CreateObject("wwDotNetBridge","V4")

loPerson = loBridge.GetPropertyEx(loFox,"Persons[0]")
? loPerson.FirstName

*** Fails: IEnumerable result doesn't work
*? loFox.persons.Count

*** Convert List<T> into an plain array
loPersonArray= loBridge.CreateArray()
loPersonArray.FromEnumerable(loFox.Persons)

FOR lnX = 0 TO loPersonArray.Count-1
	loPerson = loPersonArray.Item(lnX)
	? loPerson.FirstName + " " + loPerson.Address.Street
ENDFOR


*** Result is Person[] - automatically returned as ComArray
loPersonArray = loBridge.Invokemethod(loFox,"GetPersons")

FOR lnX = 0 TO loPersonArray.Count-1
	loPerson = loPersonArray.Item(lnX)
	? loPerson.FirstName + " " + loPerson.Address.Street
ENDFOR

? "*** Passing Arrays"

*** Demonstrate passing arrays back and fort
loPersons = loBridge.Invokemethod(loFox,"GetPersons")  && ComArray

? StringFormat("Initial Array Size: {0}", loPersons.Count)

*** Create a new Person
loNewPerson = loPersons.CreateItem()

loNewPerson.FirstName = "Billy"
loNewPerson.LastName = "Nobody"
loNewPerson.Entered = DATETIME()
loNewPerson.Address.Street = "121 Nowhere lane"

*** Add the person to the array
loPersons.AddItem(loNewPerson)

? StringFormat("After add array Size: {0}", loPersons.Count)

*** Pass the array back to .NET
loBridge.Invokemethod(loFox,"AcceptPersons",loPersons)

lnCount = loBridge.GetPropertyEx(loFox,"Persons.Count")
? StringFormat(".NET Persons Array Size after update: {0}", lnCount)



RETURN

