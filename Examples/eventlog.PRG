CLEAR

*** Accesses the Windows Event Log and writes some events
*** in the log, and then reads the log to display the last
*** 20 entries.
***
*** Demonstrates:
***   Static Method Access
***   Setting Enum Values on Methods that have overloads
***   Using ComValue to explicitly set parameter types (enum here)
***   Create a .NET Object with CreateInstance
***   Access an Array Property with ComArray

do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")

lcSource = "FoxProEvents"
lcLogType = "Application"

IF !loBridge.Invokestaticmethod("System.Diagnostics.EventLog",;
                                "SourceExists","FoxProEvents")
    loBridge.Invokestaticmethod("System.Diagnostics.EventLog",;
								"CreateEventSource",;
								"FoxProEvents","Application")
ENDIF

*** Write out default message - Information
* public static void WriteEntry(string source, string message)
loBridge.Invokestaticmethod("System.Diagnostics.EventLog",;
							"WriteEntry",lcSource,;
							"Logging from FoxPro " + TRANSFORM(DATETIME()) )



*** To use a special event log type we need to specify an enum
*** Because this method is heavily overloaded passing an enum value (ie. 1) doesn't work
*** Instead create a ComValue object and pass that instead with the enum internally assigned
LOCAL loValue as Westwind.WebConnection.ComValue
loValue = loBridge.CreateComValue()
loValue.SetEnum("System.Diagnostics.EventLogEntryType.Error")


* public static void WriteEntry(string source, string message, EventLogEntryType type, int eventID)
loBridge.Invokestaticmethod("System.Diagnostics.EventLog",;
							"WriteEntry",;
							lcSource,;
							"Logging error from FoxPro " + TRANSFORM(DATETIME()),;
							loValue, 10 )



*** Display Event Log Entries
loEventLog = loBridge.Createinstance("System.Diagnostics.EventLog","Application",".",lcSource)
*,lcSource,"Application")
*loEventLog.Source = lcSource
*loEventLog.Log = "Application"

*** Turn Eventlog Entries into a ComArray Class 
*** Indirect access automatically turns .NET array into ComArray
loEvents = loBridge.GetProperty(loEventLog,"Entries")

? "Entries: " + TRANSFORM(loEvents.Count)

lnTo =  MIN(loEvents.Count,10)
FOR lnX = loEvents.Count-1 TO loEvents.Count-lnTo STEP -1
	loEvent = loEvents.Item(lnX)
	? loEvent.message
	?
ENDFOR