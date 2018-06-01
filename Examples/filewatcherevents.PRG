CLEAR
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")

loFW = loBridge.CreateInstance("System.IO.FileSystemWatcher","C:\temp")
loFw.EnableRaisingEvents = .T.
loFw.IncludeSubDirectories = .T.

loFwHandler = CREATEOBJECT("FwEventHandler")
loSubscription = loBridge.SubscribeToEvents(loFw, loFwHandler)

DOEVENTS

lcFile = "c:\temp\test.txt"
DELETE FILE ( lcFile )  
STRTOFILE("DDD",lcFile)
STRTOFILE("FFF",lcFile)

WAIT WINDOW

loSubscription.Unsubscribe()

RETURN


DEFINE CLASS FwEventHandler as Custom

FUNCTION OnCreated(sender,ev)
? "FILE CREATED: "
?  ev.FullPath
ENDFUNC

FUNCTION OnChanged(sender,ev)
? "FILE CHANGE: "
?  ev.FullPath
ENDFUNC

FUNCTION OnDeleted(sender, ev)
? "FILE DELETED: "
?  ev.FullPath
ENDFUNC

FUNCTION OnRenamed(sender, ev)
LOCAL lcOldPath, lcPath

? "FILE RENAMED: " 
loBridge = GetwwDotnetBridge()

lcOldPath = loBridge.GetProperty(ev,"OldFullPath")
lcPath = loBridge.GetProperty(ev,"FullPath")
? lcOldPath + " -> " + lcPath

ENDFUNC

ENDDEFINE
