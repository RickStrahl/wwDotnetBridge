*** Retrieves a list of all processes running
*** on the local machine displaying status
*** information.
*** Launches a copy of notepad, then searches for
*** through process list to find it and kill it
*** 
*** Demonstrates:
***   Static Method Access
***   Creating Plain .NET Objects
***   Iterating through a ComArray Instance
CLEAR

*** Only for ShowText
DO wwutils

do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")


*** Returns a ComArray instance
loProcesses = loBridge.Invokestaticmethod("System.Diagnostics.Process","GetProcesses")

lcOutput = ""

*** Note most .NET Arrays are 0 based!
FOR lnX = 0 TO loProcesses.Count -1
   *** Access raw COM Interop objects
   loProcess = loProcesses.Item(lnX)

   lcOutput  = lcOutput + ;
   	TRANSFORM(loProcess.Id) + " -  " +loProcess.ProcessName + ;
   	"   Threads: " + TRANSFORM(loProcess.Threads.Count) + ;
   	"   Working Set: " + TRANSFORM(loProcess.WorkingSet) +;
   	CHR(13) + CHR(10)
ENDFOR

ShowText(lcOutput)

RETURN

loProcess = loBridge.CreateInstance("System.Diagnostics.Process")
loBridge.Invokemethod(loProcess,"Start","notepad.exe")
loProcess = null


WAIT WINDOW "Press any key to close NotePad"


loProcesses = loBridge.Invokestaticmethod("System.Diagnostics.Process",;
  										 "GetProcesses")
*** Track NotePads Process Id
lnNotePadId = -1

*** Find the process id of a running process
FOR lnX = 0 TO loProcesses.Count -1
   loProcess = loProcesses.Item(lnX)
   IF loProcess.ProcessName = "notepad"
      ? loProcess.ProcessName + TRANSFORM(loProcess.Id)
      lnNotePadId = loProcess.Id
   ENDIF
ENDFOR


IF lnNotePadId > -1
	loProcess = loBridge.Invokestaticmethod("System.Diagnostics.Process",;
							"GetProcessById",;
							lnNotePadId)
	loProcess.Kill()
ENDIF

