*** Demonstrates:
*** Reading an Enum Value
*** Loading a .NET CLR Type
*** Accessing an Array Property using GetPropertyEx()
*** Accessing nested properties with GetPropertyEx
CLEAR

*** Load library
DO wwDotNetBridge

*** Create instance of wwDotnetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = GetwwDotnetBridge()  && CreateObject("wwDotNetBridge")

* loBridge = GetwwDotnetCoreBridge()  && CreateObject("wwDotnetCoreBridge")

? loBridge.GetDotnetVersion()

*** Create an instance of 509Store
loStore = loBridge.CreateInstance("System.Security.Cryptography.X509Certificates.X509Store")

*** Grab a static Enum value
leReadOnly = loBridge.GetEnumvalue("System.Security.Cryptography.X509Certificates.OpenFlags","ReadOnly")

*** Use the enum value
loStore.Open(leReadOnly) &&leReadOnly)

*** Alternately you can use the enum value if known
*loStore.Open(0) &&leReadOnly)

*** Collection of Certificates
laCertificates = loStore.Certificates

*** Collections don't work over regular COM Interop
*** so use indirect access
lnCount = loBridge.GetProperty(laCertificates,"Count")

*** Loop through Certificates
FOR lnX = 1 TO lnCount -1
	*** Access collection item indirectly
	LOCAL loCertificate as System.Security.Cryptography.X509Certificates.X509Certificate2	
	loCertificate = loBridge.GetPropertyEx(loStore,"Certificates[" + TRANSFORM(lnX) + "]")
	
	IF !ISNULL(loCertificate)
		? loCertificate.FriendlyName
		? loCertificate.SerialNumber
		? loBridge.GetPropertyEx(loCertificate,"IssuerName.Name")
	ENDIF
ENDFOR

RETURN
