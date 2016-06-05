do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = GetwwDotnetBridge("V4")


loHttp = loBridge.CreateInstance("System.Net.WebClient")

*** Synchronous Call
*lohttp.DownloadFile("http://west-wind.com/files/wwclient.zip","c:\temp\wwclient.zip")

PUBLIC loCallback
loCallBack = CREATEOBJECT("HttpCallback")

loBridge.InvokeMethodAsync(loCallback, loHttp, "DownloadData","http://west-wind.com/files/wconnect.exe")

? "DONE"

RETURN

DEFINE CLASS HttpCallback as Custom

FUNCTION OnCompleted(lvResult, lcMethod)
? "Http Call completed"
? LEN(lvResult)
? lcMethod
ENDFUNC

FUNCTION OnError(lcErrorMessage, loException, lcMethod)
? lcErrorMessage
ENDFUNC

ENDDEFINE