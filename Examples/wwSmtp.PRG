*** Demonstrates how to call a custom FoxPro component
*** that has a bunch of logic that is created by
*** yourself. Optimize .NET objects for use from FoxPro
*** to keep interface simple.
***
*** Demonstrates:
*** =============
***  *  Loading an external assembly 
***  *  Create Instance
***  *  Using standard COM properties/methods
***  *  Implementing COM Event Interface (REQUIRES COM REGISTRATION)
***  *  Hooking up COM events to base class

CLEAR
*** Load helper Utilites not required by wwDotnetbridge
DO wwutils

do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")

? loBridge.LoadAssembly("InteropExamples.dll")
? loBridge.cErRORMSG

LOCAL loSmtp as Westwind.wwSmtp
*loSmtp = CREATEOBJECT("Westwind.wwSmtp")
loSmtp = loBridge.CreateInstance("Westwind.wwSmtp")

lcFile = FULLPATH("sailbig")

*loSmtp.AddAttachment_3("c:\sailbig.jpg")
loBridge.InvokeMethod(loSmtp,"AddAttachment",lcFile)

loSmtp.MailServer = "smtp.server.net:587"
loSmtp.UseSsl = .F.

loSmtp.Username = "rstrahl"
loSmtp.Password = STRTRAN(GetSystemPassword(),"0","")

loSmtp.Recipient = "John Doe<doe@doeboy.com>"
loSmtp.SenderEmail = "admin@wise-man.com"

loSmtp.Subject = "Test Message"
TEXT TO loSmtp.Message NOSHOW
<html>
	<head>
		<style>
			body { font-family: Verdana; background: cornsilk; }
		</style>
	</head>
<body>
<p>
Hello Rick,
</p>
<p>
This is a test message from <b>Southwest Fox</b>
</p>

<p>
Enjoy,
</p>
<p>
+++ Rick ---
</p>
</body>
</html>
ENDTEXT

loSmtp.contentType = "text/html"

*** Hook up the event object to wwSmtp instance
loEvents = CREATEOBJECT("wwSmtpEvents")
EVENTHANDLER(loSmtp,loEvents)

*** You can route events to another class
*!*	BINDEVENT(loEvents,"IwwSmtpEvents_SendMailMessagesObject",;
*!*	          THIS,"OnSendMailMessages")


? loSmtp.SendMailAsync()
? loSmtp.ErrorMessage

? "Code is done executing..."

RETURN


*** You can drag and drop this interface from
*** Object Browser. Make sure to change the 
*** IN clause to a COM type in the type library
DEFINE CLASS wwSmtpEvents AS session OLEPUBLIC
	IMPLEMENTS IwwSmtpEvents IN "Westwind.wwSmtp"

	PROCEDURE IwwSmtpEvents_SendError(smtp AS VARIANT) AS VOID
		? "Sending message '" + smtp.Subject + "' failed..." +;
		  smtp.ErrorMessage
	ENDPROC

	PROCEDURE IwwSmtpEvents_SendComplete(smtp AS VARIANT) AS VOID
    	? "Sending message '" + smtp.Subject + "' complete..." 
	ENDPROC
ENDDEFINE