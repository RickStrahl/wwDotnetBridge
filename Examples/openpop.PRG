CLEAR
DO wwUtils

do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge")

? loBridge.LoadAssembly("bin\OpenPop.dll")

loPop = loBridge.CreateInstance("OpenPop.Pop3.Pop3Client")
IF ISNULL(loPop)
  ? loBridge.cErrorMsg
  RETURN
ENDIF

*loPop.Connect("mail.gorge.net",587,.f.)


* loPop.Connect("mail.gorge.net",587,.f.)
? loBridge.InvokeMethod(loPop,"Connect","pop3.gorge.net",110,.f.)

? loPop.Authenticate("rstrahl",STRTRAN(GetSystemPassword(),"0",""))

lnCount =  loPop.GetMessageCount()
? StringFormat("{0} Messages",lnCount)

*** NOTE: OpenPop is 1 based because pop3 is 1 based!
** show last messages
FOR lnX = lnCount TO 1 STEP -1
   loHeader = loPop.GetMessageHeaders(lnx)
   ? loHeader.From.DisplayName
   ? "  " + loHeader.Subject
   ?
   IF lnX < lnCount - 10
      EXIT
   ENDIF
ENDFOR

*** Find the last message and display
loMsg = loPop.GetMessage(lnCount)
loMsg.Headers.Subject
    
loPart = loMsg.FindFirstHtmlVersion()
IF ISNULL(loPart)
   loPart =loMsg.FindFirstPlainTextVersion()
ENDIF

IF !ISNULL(loPart)
   ? StringFormat("Is Text: {0}",loPart.IsText)
   ShowHtml( loPart.GetBodyAsText() )
ENDIF


loPop.Disconnect()



*!*	DEFINE CLASS OpenPop as Custom 

*!*	Server = "mail.gorge.net"
*!*	Port = 110
*!*	UserName = "rstrahl"
*!*	Password = ""

*!*	FUNCTION Init()
*!*	THIS.Password = GetSystemPassword()
*!*	ENDFUNC

*!*	FUNCTION GetMessages()

*!*	LOCAL loBridge as wwDotNetBridge
*!*	loBridge = CreateObject("wwDotNetBridge")

*!*	loBridge.LoadAssembly("bin\OpenPop.dll")

*!*	loPop = loBridge.CreateInstance("OpenPop.Pop3.Pop3Client")

*!*	IF USED("TMessages")
*!*	   USE IN TMessages 
*!*	ENDIF
*!*	CREATE CURSOR TMessages( Subject c(128),Body M, From c(50), Date T)

*!*	*loPop.Connect("mail.gorge.net",587,.f.)
*!*	loException = null
*!*	*TRY
*!*		loBridge.InvokeMethod(loPop,"Connect","pop3.gorge.net",110,.f.)
*!*		loPop.Authenticate("rstrahl",STRTRAN(GetSystemPassword(),"0",""))

*!*		lnCount =  loPop.GetMessageCount()

*!*		*** NOTE: OpenPop is 1 based because pop3 is 1 based!
*!*		** show last messages
*!*		FOR lnX = lnCount TO 1 STEP -1
*!*		   loHeader = loPop.GetMessageHeaders(lnx)
*!*		   
*!*		   INSERT INTO TMessages(loHeader.
*!*		   ? loHeader.Subject
*!*		   IF lnX < lnCount - 15
*!*		      EXIT
*!*		   ENDIF
*!*		ENDFOR

*!*	    *** Find the last message and display
*!*	    loMsg = loPop.GetMessage(lnCount-1)
*!*	    loMsg.Headers.Subject
*!*	    

*!*	    loPart = loMsg.FindFirstHtmlVersion()
*!*	    IF ISNULL(loPart)
*!*	       loPart =loMsg.FindFirstPlainTextVersion()
*!*	    ENDIF
*!*	    
*!*	    IF !ISNULL(loPart)
*!*	       ? StringFormat("Is Text: {0}",loPart.IsText)
*!*	       ShowHtml( loPart.GetBodyAsText() )
*!*	    ENDIF
*!*	*CATCH TO loException
*!*	*	? loException.Message
*!*	*FINALLY
*!*	    *** Make sure you disconnect
*!*		loPop.Disconnect()
*!*	*ENDTRY
*!*	RETURN


*!*	ENDDEFINE


