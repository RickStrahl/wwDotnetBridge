clear
do wwDotNetBridge

? GetDomainFromIpAddress("127.0.0.1")
? GetDomainFromIpAddress("::1")
? GetDomainFromIpAddress("209.216.162.15")
? GetDomainFromIpAddress("209.216.162.16")
RETURN


FUNCTION GetDomainFromIpAddress(lcIpAddress)
LOCAL loBridge, llError, lcHost

loBridge = GetwwDotNetBridge()
llError = .f.
lcHost = ""

TRY
	loHost = loBridge.InvokeStaticMethod("System.Net.Dns","GetHostEntry",lcIpAddress)
	lcHost = loHost.HostName
CATCH TO loException
    llError = .t.
ENDTRY

RETURN lcHost