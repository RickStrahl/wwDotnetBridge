#INCLUDE WCONNECT.H

#DEFINE MAX_INI_BUFFERSIZE  		512
#DEFINE MAX_INI_ENUM_BUFFERSIZE 	8196

SET PROCEDURE TO wwAPI ADDITIVE

*************************************************************
DEFINE CLASS wwAPI AS Custom
*************************************************************
***    Author: Rick Strahl
***            (c) West Wind Technologies, 1997
***   Contact: (541) 386-2087  / rstrahl@west-wind.com
***  Function: Encapsulates several Windows API functions
*************************************************************

*** Custom Properties
nLastError=0
cErrorMsg = ""

FUNCTION Init
************************************************************************
* wwAPI :: Init
*********************************
***  Function: DECLARES commonly used DECLAREs so they're not redefined
***            on each call to the methods.
************************************************************************

DECLARE INTEGER GetPrivateProfileString ;
   IN WIN32API ;
   STRING cSection,;
   STRING cEntry,;
   STRING cDefault,;
   STRING @cRetVal,;
   INTEGER nSize,;
   STRING cFileName

DECLARE INTEGER GetCurrentThread ;
   IN WIN32API 
   
DECLARE INTEGER GetThreadPriority ;
   IN WIN32API ;
   INTEGER tnThreadHandle

DECLARE INTEGER SetThreadPriority ;
   IN WIN32API ;
   INTEGER tnThreadHandle,;
   INTEGER tnPriority

*** Open Registry Key
DECLARE INTEGER RegOpenKey ;
        IN Win32API ;
        INTEGER nHKey,;
        STRING cSubKey,;
        INTEGER @nHandle

*** Create a new Key
DECLARE Integer RegCreateKey ;
        IN Win32API ;
        INTEGER nHKey,;
        STRING cSubKey,;
        INTEGER @nHandle

*** Close an open Key
DECLARE Integer RegCloseKey ;
        IN Win32API ;
        INTEGER nHKey
  
ENDFUNC
* Init


FUNCTION ReadRegistryString
************************************************************************
* wwAPI :: ReadRegistryString
*********************************
***  Function: Reads a string value from the registry.
***      Pass: tnHKEY    -  HKEY value (in CGIServ.h)
***            tcSubkey  -  The Registry subkey value
***            tcEntry   -  The actual Key to retrieve
***            tlInteger -  Optional - Return an DWORD value
***            tnMaxStringSize - optional - Max size for a string (512)
***    Return: Registry String or .NULL. on not found
************************************************************************
LPARAMETERS tnHKey, tcSubkey, tcEntry, tlInteger,tnMaxStringSize
LOCAL lnRegHandle, lnResult, lnSize, lcDataBuffer, tnType

IF EMPTY(tnMaxStringSize)
   tnMaxStringSize= MAX_INI_BUFFERSIZE
ENDIF
IF EMPTY(tnHKEY)
   tnHKEY = HKEY_LOCAL_MACHINE
ENDIF   
IF VARTYPE(tnHKey) = "C"
   DO CASE
      CASE tnHKey = "HKLM"
         tnHKey = HKEY_LOCAL_MACHINE
      CASE tnHkey = "HKCU"
         tnHKey = HKEY_CURRENT_USER
      CASE tnHkey = "HKCR"
          tnHKey = HKEY_CLASSES_ROOT
      OTHERWISE 
         tnHKey = 0 
   ENDCASE
ENDIF

lnRegHandle=0

*** Open the registry key
lnResult=RegOpenKey(tnHKey,tcSubKey,@lnRegHandle)
IF lnResult#ERROR_SUCCESS
   *** Not Found
   RETURN .NULL.
ENDIF   

*** Return buffer to receive value
IF !tlInteger
*** Need to define here specifically for Return Type
*** for lpdData parameter or VFP will choke.
*** Here it's STRING.
DECLARE INTEGER RegQueryValueEx ;
        IN Win32API ;
        INTEGER nHKey,;
        STRING lpszValueName,;
        INTEGER dwReserved,;
        INTEGER @lpdwType,;
        STRING @lpbData,;
        INTEGER @lpcbData
        
	lcDataBuffer=space(tnMaxStringSize)
	lnSize=LEN(lcDataBuffer)
	lnType=REG_DWORD

	lnResult=RegQueryValueEx(lnRegHandle,tcEntry,0,@lnType,;
                         @lcDataBuffer,@lnSize)
ELSE
*** Need to define here specifically for Return Type
*** for lpdData parameter or VFP will choke. 
*** Here's it's an INTEGER
DECLARE INTEGER RegQueryValueEx ;
        IN Win32API AS RegQueryInt;
        INTEGER nHKey,;
        STRING lpszValueName,;
        INTEGER dwReserved,;
        Integer @lpdwType,;
        INTEGER @lpbData,;
        INTEGER @lpcbData

	lcDataBuffer=0
	lnSize=4
	lnType=REG_DWORD
	lnResult=RegQueryInt(lnRegHandle,tcEntry,0,@lnType,;
	                         @lcDataBuffer,@lnSize)
	IF lnResult = ERROR_SUCCESS
	   RETURN lcDataBuffer
	ELSE
       RETURN -1
	ENDIF
ENDIF
=RegCloseKey(lnRegHandle)

IF lnResult#ERROR_SUCCESS 
   *** Not Found
   RETURN .NULL.
ENDIF   

IF lnSize<2
   RETURN ""
ENDIF
   
*** Return string and strip out NULLs
RETURN SUBSTR(lcDataBuffer,1,lnSize-1)
ENDFUNC
* ReadRegistryString

************************************************************************
* Registry :: WriteRegistryString
*********************************
***  Function: Writes a string value to the registry.
***            If the value doesn't exist it's created. If the key
***            doesn't exist it is also created, but this will only
***            succeed if it's the last key on the hive.
***      Pass: tnHKEY    -  HKEY value (in WCONNECT.h)
***            tcSubkey  -  The Registry subkey value
***            tcEntry   -  The actual Key to write to
***            tcValue   -  Value to write or .NULL. to delete key
***            tlCreate  -  Create if it doesn't exist
***    Assume: Use with extreme caution!!! Blowing your registry can
***            hose your system!
***    Return: .T. or .NULL. on error
************************************************************************
FUNCTION WriteRegistryString
LPARAMETERS tnHKey, tcSubkey, tcEntry, tcValue,tlCreate
LOCAL lnRegHandle, lnResult, lnSize, lcDataBuffer, tnType

IF EMPTY(tnHKEY)
   tnHKEY = HKEY_LOCAL_MACHINE
ENDIF   
IF VARTYPE(tnHKey) = "C"
   DO CASE
      CASE tnHKey = "HKLM"
         tnHKey = HKEY_LOCAL_MACHINE
      CASE tnHkey = "HKCU"
         tnHKey = HKEY_CURRENT_USER
      CASE tnHkey = "HKCR"
          tnHKey = HKEY_CLASSES_ROOT
      OTHERWISE 
         tnHKey = 0 
   ENDCASE
ENDIF

lnRegHandle=0

lnResult=RegOpenKey(tnHKey,tcSubKey,@lnRegHandle)
IF lnResult#ERROR_SUCCESS
   IF !tlCreate
      RETURN .F.
   ELSE
      lnResult=RegCreateKey(tnHKey,tcSubKey,@lnRegHandle)
      IF lnResult#ERROR_SUCCESS
         RETURN .F.
      ENDIF  
   ENDIF
ENDIF   

*** Need to define here specifically for Return Type!
*** Here lpbData is STRING.

*** Check for .NULL. which means delete key
IF !ISNULL(tcValue)
  IF VARTYPE(tcValue) = "N"
	DECLARE INTEGER RegSetValueEx ;
	        IN Win32API ;
	        INTEGER nHKey,;
	        STRING lpszEntry,;
	        INTEGER dwReserved,;
	        INTEGER fdwType,;
	        INTEGER@ lpbData,;
	        INTEGER cbData
	  lnResult=RegSetValueEx(lnRegHandle,tcEntry,0,REG_DWORD,;
                         @tcValue,4)
  ELSE
	  DECLARE INTEGER RegSetValueEx ;
	        IN Win32API ;
	        INTEGER nHKey,;
	        STRING lpszEntry,;
	        INTEGER dwReserved,;
	        INTEGER fdwType,;
	        STRING lpbData,;
	        INTEGER cbData
	  *** Nope - write new value
	  lnSize=LEN(tcValue)
	  if lnSize=0
	     tcValue=CHR(0)
	  ENDIF

	  lnResult=RegSetValueEx(lnRegHandle,tcEntry,0,REG_SZ,;
	                         tcValue,lnSize)
  ENDIF                         
ELSE
  *** Delete a value from a key
  DECLARE INTEGER RegDeleteValue ;
          IN Win32API ;
          INTEGER nHKEY,;
          STRING cEntry

  *** DELETE THE KEY
  lnResult=RegDeleteValue(lnRegHandle,tcEntry)
ENDIF                         

=RegCloseKey(lnRegHandle)
                        
IF lnResult#ERROR_SUCCESS
   RETURN .F.
ENDIF   

RETURN .T.
ENDPROC
* WriteRegistryString

FUNCTION EnumKey
************************************************************************
* wwAPI :: EnumRegistryKey
*********************************
***  Function: Returns a registry key name based on an index
***            Allows enumeration of keys in a FOR loop. If key
***            is empty end of list is reached.
***      Pass: tnHKey    -   HKEY_ root key
***            tcSubkey  -   Subkey string
***            tnIndex   -   Index of key name to get (0 based)
***    Return: "" on error - Key name otherwise
************************************************************************
LPARAMETERS tnHKey, tcSubKey, tnIndex 
LOCAL lcSubKey, lcReturn, lnResult, lcDataBuffer

lnRegHandle=0

*** Open the registry key
lnResult=RegOpenKey(tnHKey,tcSubKey,@lnRegHandle)
IF lnResult#ERROR_SUCCESS
   *** Not Found
   RETURN .NULL.
ENDIF   

DECLARE Integer RegEnumKey ;
  IN WIN32API ;
  INTEGER nHKey, ;
  INTEGER nIndex, ;
  STRING @cSubkey, ;  
  INTEGER nSize

lcDataBuffer=SPACE(MAX_INI_BUFFERSIZE)
lnSize=MAX_INI_BUFFERSIZE
lnResult=RegENumKey(lnRegHandle, tnIndex, @lcDataBuffer, lnSize)

=RegCloseKey(lnRegHandle)

IF lnResult#ERROR_SUCCESS 
   *** Not Found
   RETURN .NULL.
ENDIF   

RETURN TRIM(CHRTRAN(lcDataBuffer,CHR(0),""))
ENDFUNC
* EnumRegistryKey


FUNCTION GetProfileString
************************************************************************
* wwAPI :: GetProfileString
***************************
***  Modified: 09/26/95
***  Function: Read Profile String information from a given
***            text file using Windows INI formatting conventions
***      Pass: pcFileName   -    Name of INI file
***            pcSection    -    [Section] in the INI file ("Drivers")
***            pcEntry      -    Entry to retrieve ("Wave")
***                              If this value is a null string
***                              all values for the section are
***                              retrieved seperated by CHR(13)s
***    Return: Value(s) or .NULL. if not found
************************************************************************
LPARAMETERS pcFileName,pcSection,pcEntry, pnBufferSize
LOCAL lcIniValue, lnResult

*** Initialize buffer for result
lcIniValue=SPACE(IIF( vartype(pnBufferSize)="N",pnBufferSize,MAX_INI_BUFFERSIZE) )

lnResult=GetPrivateProfileString(pcSection,pcEntry,"*None*",;
   @lcIniValue,LEN(lcIniValue),pcFileName)

*** Strip out Nulls
IF VARTYPE(pcEntry)="N" AND pcEntry=0
   *** 0 was passed to get all entry labels
   *** Seperate all of the values with a Carriage Return
   lcIniValue=TRIM(CHRTRAN(lcIniValue,CHR(0),CHR(13)) )
ELSE
   *** Individual Entry
   lcIniValue=SUBSTR(lcIniValue,1,lnResult)
ENDIF

*** On error the result contains "*None"
IF lcIniValue="*None*"
   lcIniValue=.NULL.
ENDIF

RETURN lcIniValue
ENDFUNC
* GetProfileString

************************************************************************
* wwAPI :: GetProfileSections
*********************************
***  Function: Retrieves all sections of an INI File
***      Pass: @laSections   -   Empty array to receive sections
***            lcIniFile     -   Name of the INI file
***            lnBufSize     -   Size of result buffer (optional)
***    Return: Count of Sections  
************************************************************************
FUNCTION aProfileSections
LPARAMETERS laSections, lcIniFile
LOCAL lnBufsize, lcBuffer, lnSize, lnResult, lnCount

lnBufsize=IIF(EMPTY(lnBufsize),16484,lnBufsize)

DECLARE INTEGER GetPrivateProfileSectionNames ;
   IN WIN32API ;
   STRING @lpzReturnBuffer,;
   INTEGER nSize,;
   STRING lpFileName
   
lcBuffer = SPACE(lnBufSize)
lnSize = lEN(lcBuffer)   
lnResult = GetPrivateProfileSectionNames(@lcBuffer,lnSize,lcIniFile)
IF lnResult < 3
   RETURN 0
ENDIF

lnCount = aParseString(@laSections,TRIM(lcBuffer),CHR(0))
lnCount = lnCount - 2
IF lnCount > 0
  DIMENSION laSections[lnCount]
ENDIF

RETURN lnCount
ENDFUNC
* wwAPI :: aProfileSections

************************************************************************
* wwAPI :: WriteProfileString
*********************************
***  Function: Writes a value back to an INI file
***      Pass: pcFileName    -   Name of the file to write to
***            pcSection     -   Profile Section
***            pcKey         -   The key to write to
***            pcValue       -   The value to write
***    Return: .T. or .F.
************************************************************************
FUNCTION WriteProfileString
LPARAMETERS pcFileName,pcSection,pcEntry,pcValue

   DECLARE INTEGER WritePrivateProfileString ;
      IN WIN32API ;
      STRING cSection,STRING cEntry,STRING cValue,;
      STRING cFileName

   lnRetVal=WritePrivateProfileString(pcSection,pcEntry,pcValue,pcFileName)

   if lnRetval=1
      RETURN .t.
   endif
   
   RETURN .f.
ENDFUNC
* WriteProfileString

FUNCTION GetTempPath
************************************************************************
* wwAPI :: GetTempPath
***********************
***  Function: Returns the OS temporary files path
***    Return: Temp file path with trailing "\"
************************************************************************
LOCAL lcPath, lnResult

*** API Definition:
*** ---------------
*** DWORD GetTempPath(cchBuffer, lpszTempPath)
***
*** DWORD cchBuffer;	/* size, in characters, of the buffer	*/
*** LPTSTR lpszTempPath;	/* address of buffer for temp. path name	*/
DECLARE INTEGER GetTempPath ;
   IN WIN32API AS GetTPath ;
   INTEGER nBufSize, ;
   STRING @cPathName

lcPath=SPACE(256)
lnSize=LEN(lcPath)

lnResult=GetTPath(lnSize,@lcPath)

IF lnResult=0
   lcPath=""
ELSE
   lcPath=SUBSTR(lcPath,1,lnResult)
ENDIF

RETURN lcPath
ENDFUNC
* eop GetTempPath


FUNCTION MessageBeep
************************************************************************
* wwAPI :: MessageBeep
**********************
***  Function: MessageBeep API call runs system sounds
***      Pass: lnSound   -   Uses FoxPro.h MB_ICONxxxxx values
***    Return: nothing
************************************************************************
LPARAMETERS lnSound
DECLARE INTEGER MessageBeep ;
   IN WIN32API AS MsgBeep ;
   INTEGER nSound
=MsgBeep(lnSound)
ENDFUNC
* MessageBeep

************************************************************************
FUNCTION MapDrive()
****************************************
***  Function: Maps a network path
***    Assume:
***      Pass: lcNetPath -  \\servername\share  (\\rasvist\f$)
***            lcShareName =  z:
***            lcPassword - optional
***    Return:
************************************************************************
LPARAMETERS lcNetPath, lcShareName, lcPassword

IF EMPTY(lcPassword)
  lcPassword = CHR(0)
ENDIF

DECLARE INTEGER WNetAddConnection IN WIN32API ;
    string lpRemoteName, string lpPassWord, string lpLocalName
    
lnError =  WNetAddConnection(lcNetPath,lcPassword, lcShareName)
IF lnError # 0
	this.cErrorMsg = this.GetSystemErrorMsg(lnError)
	RETURN .F.
ENDIF

RETURN .T.
ENDFUNC
*  wwAPI ::  MapPath

************************************************************************
FUNCTION DisconnectDrive
****************************************
***  Function: Disconnects a network drive mapping
***    Assume:
***      Pass:
***    Return:
************************************************************************
LPARAMETERS lcShareName
LOCAL lnError

DECLARE INTEGER WNetCancelConnection in Win32API ;
   string lpName, INTEGER bForce
   
lnError = WNetCancelConnection(lcShareName,1)   
IF lnError # 0
   this.cErrorMsg = this.GetSystemErrorMsg(lnError)
   RETURN .F.
ENDIF

RETURN .T.   
ENDFUNC
*  wwAPI ::  UnmapDrive


FUNCTION GetEXEFile
************************************************************************
* wwAPI :: GetEXEFileName
*********************************
***  Function: Returns the Module name of the EXE file that started
***            the current application. Unlike Application.Filename
***            this function correctly returns the name of the EXE file
***            for Automation servers too!
***    Return: Filename or ""  (VFP.EXE is returned in Dev Version)
************************************************************************
DECLARE integer GetModuleFileName ;
   IN WIN32API ;
   integer hinst,;
   string @lpszFilename,;
   integer @cbFileName
   
lcFilename=space(256)
lnBytes=255   

=GetModuleFileName(0,@lcFileName,@lnBytes)

lnBytes=AT(CHR(0),lcFileName)
IF lnBytes > 1
  lcFileName=SUBSTR(lcFileName,1,lnBytes-1)
ELSE
  lcFileName=""
ENDIF       

RETURN lcFileName
ENDFUNC
* GetEXEFileName


************************************************************************
* WinApi :: ShellExecute
*********************************
***    Author: Rick Strahl, West Wind Technologies
***            http://www.west-wind.com/ 
***  Function: Opens a file in the application that it's
***            associated with.
***      Pass: lcFileName  -  Name of the file to open
***            lcWorkDir   -  Working directory
***            lcOperation -  
***    Return: 2  - Bad Association (invalid URL)
***            31 - No application association
***            29 - Failure to load application
***            30 - Application is busy 
***
***            Values over 32 indicate success
***            and return an instance handle for
***            the application started (the browser) 
************************************************************************
***         FUNCTION ShellExecute
***         LPARAMETERS lcFileName, lcWorkDir, lcOperation
***         
***         lcWorkDir=IIF(type("lcWorkDir")="C",lcWorkDir,"")
***         lcOperation=IIF(type("lcOperation")="C",lcOperation,"Open")
***         
***         DECLARE INTEGER ShellExecute ;
***             IN SHELL32.DLL ;
***             INTEGER nWinHandle,;
***             STRING cOperation,;   
***             STRING cFileName,;
***             STRING cParameters,;
***             STRING cDirectory,;
***             INTEGER nShowWindow
***            
***         RETURN ShellExecute(0,lcOperation,lcFilename,"",lcWorkDir,1)
***         ENDFUNC
***         * ShellExecute

************************************************************************
* wwAPI :: CopyFile
*********************************
***  Function: Copies File. Faster than Fox Copy and handles
***            errors internally.
***      Pass: tcSource -  Source File
***            tcTarget -  Target File
***            tnFlag   -  0* overwrite, 1 don't overwrite
***    Return: .T. or .F.
************************************************************************
FUNCTION CopyFile
LPARAMETERS lcSource, lcTarget,nFlag
LOCAL lnRetVal 

*** Copy File and overwrite
nFlag=IIF(vartype(nFlag)="N",nFlag,0)

DECLARE INTEGER CopyFile ;
   IN WIN32API ;
   STRING @cSource,;
   STRING @cTarget,;
   INTEGER nFlag

lnRetVal=CopyFile(@lcSource,@lcTarget,nFlag)

RETURN IIF(lnRetVal=0,.F.,.T.)
ENDPROC
* CopyFile


FUNCTION GetUserName


DECLARE INTEGER GetUserName ;
     IN WIN32API ;
     STRING@ cComputerName,;
     INTEGER@ nSize

lcComputer=SPACE(80)
lnSize=80

=GetUserName(@lcComputer,@lnSize)
IF lnSize < 2
   RETURN ""
ENDIF   

RETURN SUBSTR(lcComputer,1,lnSize-1)

FUNCTION GetComputerName
************************************************************************
* wwAPI :: GetComputerName
*********************************
***  Function: Returns the name of the current machine
***    Return: Name of the computer
************************************************************************

DECLARE INTEGER GetComputerName ;
     IN WIN32API ;
     STRING@ cComputerName,;
     INTEGER@ nSize

lcComputer=SPACE(80)
lnSize=80

=GetComputername(@lcComputer,@lnSize)
IF lnSize < 2
   RETURN ""
ENDIF   

RETURN SUBSTR(lcComputer,1,lnSize)
ENDFUNC
* GetComputerName


FUNCTION LogonUser
************************************************************************
* wwAPI :: LogonUser
*********************************
***  Function: Check whether a username and password is valid
***    Assume: Account checking must have admin rights
***      Pass: Username, Password and optionally a server
***            lcServer   -  The server name (or . for local)
***            lnToken    -  Pass in a @lnToken if you want to get
***                          the impersonation token provided 
***    Return: .T. or .F.
************************************************************************
LPARAMETERS lcUsername, lcPassword, lcServer, lnToken
LOCAL lnResult,llTokenPassed

IF EMPTY(lcUsername)
   RETURN .F.
ENDIF
IF EMPTY(lcPassword)
   lcPassword = ""
ENDIF
IF EMPTY(lcServer)
   lcServer = "."
ENDIF         

IF VARTYPE(lnToken) = "N"
  llTokenPassed = .T.
ENDIF

#define LOGON32_LOGON_INTERACTIVE   2
#define LOGON32_LOGON_NETWORK       3
#define LOGON32_LOGON_BATCH         4
#define LOGON32_LOGON_SERVICE       5

#define LOGON32_PROVIDER_DEFAULT    0

DECLARE INTEGER LogonUser in WIN32API ;
       String lcUser,;
       String lcServer,;
       String lcPassword,;
       INTEGER dwLogonType,;
       Integer dwProvider,;
       Integer @dwToken
       
lnToken = 0
lnResult = LogonUser(lcUsername,lcServer,lcPassword,;
                     LOGON32_LOGON_NETWORK,LOGON32_PROVIDER_DEFAULT,@lnToken) 

IF !llTokenPassed
	DECLARE INTEGER CloseHandle IN WIN32API INTEGER
	CloseHandle(lnToken)
ENDIF

RETURN IIF(lnResult=1,.T.,.F.)
ENDFUNC
* wwAPI :: LogonUser


************************************************************************
*  ImpersonateUser
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ImpersonateUser(lcUsername, lcPassword)
LOCAL lnToken, lnResult

lnToken = 0
IF !this.LogonUser(lcUsername,lcPassword,.f.,@lnToken)
	RETURN .F.
ENDIF
	
DECLARE integer ImpersonateLoggedOnUser ;
	IN WIN32API ;
	integer hToken
	
lnResult = ImpersonateLoggedOnUser(lnToken)

DECLARE INTEGER CloseHandle IN WIN32API INTEGER
CloseHandle(lnToken)

RETURN IIF(lnResult=1,.T.,.F.)
ENDFUNC
*   ImpersonateUser


************************************************************************
*  RevertToSelf
****************************************
***  Function: Reverts the account to the base account that launched
***            the application.
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION RevertToSelf()
DECLARE integer RevertToSelf IN Win32API
RETURN IIF( RevertToSelf() = 1,.T.,.F.)
ENDFUNC
*   RevertToSelf

FUNCTION GetSystemDir
************************************************************************
* wwAPI :: GetSystemDir
*********************************
***  Function: Returns the Windows System directory path
***      Pass: llWindowsDir - Optional: Retrieve the Windows dir
***    Return: Windows System directory or "" if failed
************************************************************************
LPARAMETER llWindowsDir
LOCAL lcPath, lnSize

lcPath=SPACE(256)

IF !llWindowsDir
	DECLARE INTEGER GetSystemDirectory ;
	   IN Win32API ;
	   STRING  @pszSysPath,;
	   INTEGER cchSysPath
	lnsize=GetSystemDirectory(@lcPath,256) 
ELSE
	DECLARE INTEGER GetWindowsDirectory ;
	   IN Win32API ;
	   STRING  @pszSysPath,;
	   INTEGER cchSysPath
	lnsize=GetWindowsDirectory(@lcPath,256) 
ENDIF 

if lnSize > 0
   RETURN SUBSTR(lcPath,1,lnSize) + "\"
ENDIF
   
RETURN ""
ENDFUNC
* GetSystemDir

   

FUNCTION GetCurrentThread
************************************************************************
* wwAPI :: GetCurrentThread
*********************************
***  Function: Returns handle to the current Process/Thread
***    Return: Process Handle or 0
************************************************************************
RETURN GetCurrentThread()
ENDFUNC
* GetProcess

************************************************************************
* wwAPI :: GetThreadPriority
*********************************
***  Function: Gets the current Priority setting of the thread.
***            Use to save and reset priority when bumping it up.
***      Pass: tnThreadHandle
************************************************************************
FUNCTION GetThreadPriority
LPARAMETER tnThreadHandle
RETURN GetThreadPriority(tnThreadHandle)
ENDFUNC
* GetThreadPriority

FUNCTION SetThreadPriority
************************************************************************
* wwAPI :: SetThreadPriority
*********************************
***  Function: Sets a thread process priority. Can dramatically
***            increase performance of a task.
***      Pass: tnThreadHandle
***            tnPriority         0 - Normal
***                               1 - Above Normal
***                               2 - Highest Priority
***                              15 - Time Critical
***                              31 - Real Time (doesn't work w/ Win95)
************************************************************************
LPARAMETER tnThreadHandle,tnPriority
RETURN SetThreadPriority(tnThreadHandle,tnPriority)
ENDFUNC
* GetThreadPriority


FUNCTION PlayWave
************************************************************************
* wwapi :: PlayWave
*******************
***     Class: WinAPI
***  Function: Plays the Wave File or WIN.INI
***            [Sounds] Entry specified in the
***            parameter. If the .WAV file or
***            System Sound can't be found,
***            SystemDefault beep is played
***    Assume: Runs only under Windows
***            uses MMSYSTEM.DLL  (Win 3.1)
***                 WINMM.DLL  (32 bit Win)
***      Pass: pcWaveFile - Full path of Wave file
***                         or System Sound Entry
***            pnPlayType - 1 - sound plays in background (default)
***                         0 - sound plays - app waits
***                         2 - No default sound if file doesn't exist
***                         4 - Kill currently playing sound 
***                         8 - Continous  
***                         Values can be added together for combinations
***  Examples:
***    do PlayWav with "SystemQuestion"
***    do PlayWav with "C:\Windows\System\Ding.wav"
***    if PlayWav("SystemAsterisk")
***
***    Return: .t. if Wave was played .f. otherwise
*************************************************************************
LPARAMETER pcWaveFile,pnPlayType
LOCAL lhPlaySnd,llRetVal

pnPlayType=IIF(TYPE("pnPlayType")="N",pnPlayType,1)

llRetVal=.f.

DECLARE INTEGER PlaySound ;
   IN WINMM.dll  ;
   STRING cWave, INTEGER nModule, INTEGER nType

IF PlaySound(pcWaveFile,0,pnPlayType)=1
   llRetVal=.t.
ENDIF

RETURN llRetVal
ENDFUNC
*EOF PLAYWAV


FUNCTION CreateGUID
************************************************************************
* wwapi::CreateGUID
********************
***    Author: Rick Strahl, West Wind Technologies
***            http://www.west-wind.com/
***  Modified: 01/26/98
***  Function: Creates a globally unique identifier using Win32
***            COM services. The vlaue is guaranteed to be unique
***    Format: {9F47F480-9641-11D1-A3D0-00600889F23B}
***    Return: GUID as a string or "" if the function failed 
*************************************************************************
LPARAMETERS llRaw
LOCAL lcStruc_GUID, lcGUID, lnSize

DECLARE INTEGER CoCreateGuid ;
  IN Ole32.dll ;
  STRING @lcGUIDStruc
  
DECLARE INTEGER StringFromGUID2 ;
  IN Ole32.dll ;
  STRING cGUIDStruc, ;
  STRING @cGUID, ;
  LONG nSize
  
*** Simulate GUID strcuture with a string
lcStruc_GUID = REPLICATE(" ",16) 
lcGUID = REPLICATE(" ",80)
lnSize = LEN(lcGUID) / 2
IF CoCreateGuid(@lcStruc_GUID) # 0
   RETURN ""
ENDIF

IF llRaw
   RETURN lcStruc_GUID
ENDIF   

*** Now convert the structure to the GUID string
IF StringFromGUID2(lcStruc_GUID,@lcGuid,lnSize) = 0
  RETURN ""
ENDIF

*** String is UniCode so we must convert to ANSI
RETURN  StrConv(LEFT(lcGUID,76),6)
* Eof CreateGUID

FUNCTION Sleep(lnMilliSecs)
************************************************************************
* wwAPI :: Sleep
*********************************
***  Function: Puts the computer into idle state. More efficient and
***            no keyboard interface than Inkey()
***      Pass: tnMilliseconds
***    Return: nothing
************************************************************************

lnMillisecs=IIF(type("lnMillisecs")="N",lnMillisecs,0)

DECLARE Sleep ;
  IN WIN32API ;
  INTEGER nMillisecs
 	
=Sleep(lnMilliSecs) 	
ENDFUNC
* Sleep

************************************************************************
* wwAPI :: GetLastError
*********************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetLastError
DECLARE INTEGER GetLastError IN Win32API 
RETURN GetLastError()
ENDFUNC
* wwAPI :: GetLastError

************************************************************************
* wwAPI :: GetSystemErrorMsg
*********************************
***  Function: Returns the Message text for a Win32API error code.
***      Pass: lnErrorNo  -  WIN32 Error Code
***    Return: Error Message or "" if not found
************************************************************************
FUNCTION GetSystemErrorMsg
LPARAMETERS lnErrorNo,lcDLL
LOCAL szMsgBuffer,lnSize

szMsgBuffer=SPACE(500)

DECLARE INTEGER FormatMessage ;
     IN WIN32API ;
     INTEGER dwFlags ,;
     STRING lpvSource,;
     INTEGER dwMsgId,;
     INTEGER dwLangId,;
     STRING @lpBuffer,;
     INTEGER nSize,;
     INTEGER  Arguments

lnSize=FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM,0,lnErrorNo,;
                     0,@szMsgBuffer,LEN(szMsgBuffer),0)

IF LEN(szMsgBUffer) > 1
  szMsgBuffer=SUBSTR(szMsgBuffer,1, lnSize-1 )
ELSE
  szMsgBuffer=""  
ENDIF
    		   
RETURN szMsgBuffer


ENDDEFINE
*EOC wwAPI


************************************************************************
* wwAPI :: GetSpecialFolder
****************************************
***  Function:
***    Assume: 0x002b -  Common Files
***            0x0026 -  Program Files
***            
***            
***      Pass:
***    Return:
************************************************************************
FUNCTION GetSpecialFolder(lnFolder)

IF VARTYPE(lnFolder) = "C"
   DO CASE
      *** MSDN -  CSIDL flag translates
      CASE lnFolder = "Program Files Common"
         lnFolder = 0x002B
      CASE lnFolder = "Program Files"  && x86
         lnFolder = 0x0026
      CASE lnFolder = "Program Files 64Bit"   && 64 bit folder
         lnFolder = 0x002a
      CASE lnFolder = "Documents Common"
         lnFolder = 0x002E
      CASE  lnFolder == "Documents" OR lnFolder = "Documents User" OR lnFolder = "My Documents"
         lnFolder = 0x0005
      CASE lnFolder = "Send To"
         lnFolder = 0x0009
      CASE lnFolder = "My Computer"
         lnFolder = 0x0011
      CASE lnFolder = "Desktop"
         lnFolder = 0
      CASE lnFolder == "Application Data"
         lnFolder = 0x001A
      CASE lnFolder == "Application Data Common"
         lnFolder = 0x0023
   ENDCASE
ENDIF


DECLARE INTEGER SHGetFolderPath IN Shell32.dll ;
      INTEGER Hwnd, INTEGER nFolder, INTEGER Token, INTEGER Flags, STRING @cPath
      
lcOutput = repl(CHR(0),256)
lnResult = SHGetFolderPath(_VFP.hWnd,lnFolder,0,0,@lcOutput)
IF lnResult = 0
   lcOutput = STRTRAN(lcOutput,CHR(0),"") + "\"
ELSE
   lcOutput = ""
ENDIF

RETURN lcOutput
ENDFUNC
*  wwAPI :: GetSpecialFolder

************************************************************************
* wwAPI :: CreateShortcut
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION CreateShortcut(lcShortCut,lcDescription, lcTarget,lcArguments,lcStartFolder,lcIcon)

IF !ISCOMOBJECT("wscript.Shell")
   RETURN .f.
ENDIF

   
llError = .f.   
*TRY 
loScript = create("wscript.Shell")
loSc = loScript.createShortCut(lcShortCut)
loSC.Description = lcDescription
loSC.TargetPath = lcTarget

IF !EMPTY(lcArguments)
   loSC.Arguments = lcArguments
ENDIF
IF !EMPTY(lcIcon)
loSC.IconLocation = lcIcon
ENDIF

IF EMPTY(lcStartFolder)
   loSC.WorkingDirectory = JUSTPATH(lcTarget)
ELSE
   loSC.WorkingDirectory = lcStartFolder
ENDIF

loSC.Save()
*CATCH
*   llError = .t.
*ENDTRY

RETURN !llError  
ENDFUNC
*  wwAPI :: CreateShortcut




************************************************************************
*  GetUtcTime
****************************************
***  Function: Returns UTC time from local time
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetUtcTime(ltTime)

IF EMPTY(ltTime)
    ltTime = DATETIME()
ENDIF

*** Adjust the timezone offset
RETURN ltTime + (GetTimeZone() * 60)	
ENDFUNC
*   GetUtcTime

************************************************************************
*  FromUtcTime
****************************************
***  Function: Returns local time from UTC Time
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION FromUtcTime(ltTime)
RETURN ltTime - (GetTimeZone() * 60)
ENDFUNC
*   FromUtcTime

************************************************************************
FUNCTION GetTimeZone
*********************************
***  Function: Returns the TimeZone offset from GMT including
***            daylight savings. Result is returned in minutes.
************************************************************************
PUBLIC __TimeZone

*** Cache the timezone so this is fast
IF VARTYPE(__TimeZone) = "N"
   RETURN __TimeZone
ENDIF

DECLARE integer GetTimeZoneInformation IN Win32API ;
   STRING @ TimeZoneStruct
   
lcTZ = SPACE(256)

lnDayLightSavings = GetTimeZoneInformation(@lcTZ)
lnOffset = CharToBin(SUBSTR(lcTZ,1,4),.T.)

*** Subtract an hour if daylight savings is active
IF lnDaylightSavings = 2
   lnOffset = lnOffset - 60
ENDIF

__TimeZone = lnOffset
	
RETURN lnOffSet

************************************************************************
FUNCTION CharToBin(lcBinString,llSigned)
****************************************
***  Function: Binary Numeric conversion routine. 
***            Converts DWORD or Unsigned Integer string
***            to Fox numeric integer value.
***      Pass: lcBinString -  String that contains the binary data 
***            llSigned    -  if .T. uses signed conversion
***                           otherwise value is unsigned (DWORD)
***    Return: Fox number
************************************************************************
LOCAL m.i, lnWord

lnWord = 0
FOR m.i = 1 TO LEN(lcBinString)
 lnWord = lnWord + (ASC(SUBSTR(lcBinString, m.i, 1)) * (2 ^ (8 * (m.i - 1))))
ENDFOR

IF llSigned AND lnWord > 0x80000000
  lnWord = lnWord - 1 - 0xFFFFFFFF
ENDIF

RETURN lnWord
*  wwAPI :: CharToBin

************************************************************************
FUNCTION BinToChar(lnValue)
****************************************
***  Function: Creates a DWORD value from a number
***      Pass: lnValue - VFP numeric integer (unsigned)
***    Return: binary string
************************************************************************
Local byte(4)
If lnValue < 0
    lnValue = lnValue + 4294967296
EndIf
byte(1) = lnValue % 256
byte(2) = BitRShift(lnValue, 8) % 256
byte(3) = BitRShift(lnValue, 16) % 256
byte(4) = BitRShift(lnValue, 24) % 256
RETURN Chr(byte(1))+Chr(byte(2))+Chr(byte(3))+Chr(byte(4))
*  wwAPI :: BinToChar

************************************************************************
FUNCTION BinToWordChar(lnValue)
****************************************
***  Function: Creates a DWORD value from a number
***      Pass: lnValue - VFP numeric integer (unsigned)
***    Return: binary string
************************************************************************
RETURN Chr(MOD(m.lnValue,256)) + CHR(INT(m.lnValue/256))


************************************************************************
* wwAPI :: FindWindow
****************************************
***  Function: Returns a Window Handle for a window on the desktop
************************************************************************
FUNCTION FindWindow(lcTitle)
DECLARE INTEGER FindWindow IN WIN32API AS __FindWindow integer Handle, STRING Title
RETURN __FindWindow(0,lcTitle)
ENDFUNC


#IF wwVFPVersion > 7
************************************************************************
* wwAPI :: HashMD5
****************************************
***  Function: retrieved from the FoxWiki
*** 		   http://fox.wikis.com/wc.dll?fox~vfpmd5hashfunction
***    Assume: Self standing function - not part of wwAPI class
***      Pass: Data to encrypt
***    Return: 
************************************************************************
FUNCTION HashMD5(tcData)

*** #include "c:\program files\microsoft visual foxpro 8\ffc\wincrypt.h"
#DEFINE dnPROV_RSA_FULL           1
#DEFINE dnCRYPT_VERIFYCONTEXT     0xF0000000

#DEFINE dnALG_CLASS_HASH         BITLSHIFT(4,13)
#DEFINE dnALG_TYPE_ANY 		 0
#DEFINE dnALG_SID_MD5           3
#DEFINE dnCALG_MD5        BITOR(BITOR(dnALG_CLASS_HASH,dnALG_TYPE_ANY),dnALG_SID_MD5)

#DEFINE dnHP_HASHVAL              0x0002  && Hash value

LOCAL lnStatus, lnErr, lhProv, lhHashObject, lnDataSize, lcHashValue, lnHashSize
lhProv = 0
lhHashObject = 0
lnDataSize = LEN(tcData)
lcHashValue = REPLICATE(CHR(0), 16)
lnHashSize = LEN(lcHashValue)


DECLARE INTEGER GetLastError ;
   IN win32api AS GetLastError

DECLARE INTEGER CryptAcquireContextA ;
   IN WIN32API AS CryptAcquireContext ;
   INTEGER @lhProvHandle, ;
   STRING cContainer, ;
   STRING cProvider, ;
   INTEGER nProvType, ;
   INTEGER nFlags

* load a crypto provider
lnStatus = CryptAcquireContext(@lhProv, 0, 0, dnPROV_RSA_FULL, dnCRYPT_VERIFYCONTEXT)
IF lnStatus = 0
   THROW GetLastError()
ENDIF

DECLARE INTEGER CryptCreateHash ;
   IN WIN32API AS CryptCreateHash ;
   INTEGER hProviderHandle, ;
   INTEGER nALG_ID, ;
   INTEGER hKeyhandle, ;
   INTEGER nFlags, ;
   INTEGER @hCryptHashHandle

* create a hash object that uses MD5 algorithm
lnStatus = CryptCreateHash(lhProv, dnCALG_MD5, 0, 0, @lhHashObject)
IF lnStatus = 0
   THROW GetLastError()
ENDIF

DECLARE INTEGER CryptHashData ;
   IN WIN32API AS CryptHashData ;
   INTEGER hHashHandle, ;
   STRING @cData, ;
   INTEGER nDataLen, ;
   INTEGER nFlags

* add the input data to the hash object
lnStatus = CryptHashData(lhHashObject, tcData, lnDataSize, 0)
IF lnStatus = 0
   THROW GetLastError()
ENDIF


DECLARE INTEGER CryptGetHashParam ;
   IN WIN32API AS CryptGetHashParam ;
   INTEGER hHashHandle, ;
   INTEGER nParam, ;
   STRING @cHashValue, ;
   INTEGER @nHashSize, ;
   INTEGER nFlags

* retrieve the hash value, if caller did not provide enough storage (16 bytes for MD5)
* this will fail with dnERROR_MORE_DATA and lnHashSize will contain needed storage size
lnStatus = CryptGetHashParam(lhHashObject, dnHP_HASHVAL, @lcHashValue, @lnHashSize, 0)
IF lnStatus = 0
   THROW GetLastError()
ENDIF


DECLARE INTEGER CryptDestroyHash ;
   IN WIN32API AS CryptDestroyHash;
   INTEGER hKeyHandle

*** free the hash object
lnStatus = CryptDestroyHash(lhHashObject)
IF lnStatus = 0
   THROW GetLastError()
ENDIF


DECLARE INTEGER CryptReleaseContext ;
   IN WIN32API AS CryptReleaseContext ;
   INTEGER hProvHandle, ;
   INTEGER nReserved

*** release the crypto provider
lnStatus = CryptReleaseContext(lhProv, 0)
IF lnStatus = 0
   THROW GetLastError()
ENDIF

RETURN lcHashValue
ENDFUNC
* HashMD5
#ENDIF




************************************************************************
FUNCTION ResizeImage(lcSource,lcTarget,lnWidth,lnHeight,lnCompression)
****************************************
***  Function: Creates a Thumbnail image from a file into another file
***    Assume:
***      Pass:
***    Return:
************************************************************************

IF EMPTY(lnCompression)
  lnCompression = -1
ENDIF
IF lnCompression > 100 OR lnCompression < -1
   lnCompression = -1
ENDIF

DECLARE INTEGER ResizeImage IN wwImaging.dll AS _ResizeImage;
   STRING lcSource, STRING lcTarget, INTEGER lnWidth, INTEGER lnHeight, INTEGER lnCompression
   
RETURN (IIF(_ResizeImage(STRCONV(FULLPATH(lcSource)+CHR(0),5),;
                             STRCONV(LOWER(FULLPATH(lcTarget))+CHR(0),5),;
                             lnWidth,lnHeight,lnCompression)=1,.T.,.F.))

************************************************************************
FUNCTION CopyImage(lcSource,lcTarget)
****************************************
***  Function: Copies an image from one format to another
***    Assume:
***      Pass:
***    Return:
************************************************************************

IF LOWER(JUSTEXT(lcTarget)) = "gif"  
   DECLARE INTEGER SaveImageToGif IN wwImaging.dll as _SaveImageAsGif ;
                   STRING lcSource, STRING lcTarget
    RETURN (IIF(_SaveImageAsGif(STRCONV(FULLPATH(lcSource)+CHR(0),5),;
                           STRCONV(FULLPATH(lcTarget)+CHR(0),5))=1,.T.,.F.))
ELSE
    DECLARE INTEGER CopyImageEx IN wwImaging.dll AS _CopyImage ;
                    STRING lcSource, STRING lcTarget

    RETURN (IIF(_CopyImage(STRCONV(FULLPATH(lcSource)+CHR(0),5),;
                           STRCONV(LOWER(FULLPATH(lcTarget))+CHR(0),5))=1,.T.,.F.))
ENDIF

                             
************************************************************************
FUNCTION CreateThumbNail(lcSource,lcTarget,lnWidth,lnHeight)
****************************************
***  Function: Creates a Thumbnail image from a file into another file
***    Assume:
***      Pass:
***    Return:
************************************************************************

DECLARE INTEGER CreateThumbnail IN wwImaging.dll AS _CreateThumbnail;
   STRING lcSource, STRING lcTarget, INTEGER lnWidth, INTEGER lnHeight
   
RETURN (IIF(_CreateThumbnail(STRCONV(FULLPATH(lcSource)+CHR(0),5),;
                             STRCONV(FULLPATH(lcTarget)+CHR(0),5),;
                             lnWidth,lnHeight)=1,.T.,.F.))

************************************************************************
FUNCTION GetImageInfo(lcImage,lnWidth,lnHeight,lnResolution)
****************************************
***  Function: Returns Width, Height and Resolution of an image
***    Assume:
***      Pass: Pass the last 3 parameters in by Reference
***    Return:
************************************************************************

DECLARE INTEGER GetImageInfo IN wwImaging.dll AS _GetImageInfo;
   STRING lcSource, INTEGER@ lnWidth, INTEGER@ lnHeight, INTEGER@ lnResolution
   
lnWidth = 0
lnHeight = 0
lnResolution = 0

RETURN IIF(;
 _GetImageInfo(STRCONV(FULLPATH(lcImage) +CHR(0) ,5),;
              @lnWidth,@lnHeight,@lnResolution) = 1,.T.,.F.)

ENDFUNC
* GetImageInfo

************************************************************************
FUNCTION RotateImage(lcImage,lnFlipType)
****************************************
***  Function: Returns Width, Height and Resolution of an image
***    Assume:
***      Pass: lcImage    - Image file to convert in place
***            lnFlipType - Type of rotation or flip to perform
***                         1  -   Rotate 90 degrees
***                         2  -   Rotate 180 degrees
***                         3  -   Rotate 270 degrees
***                         4  -   Flip Image (mirror image)
***                         5  -   Flip Image and Rotate 90 degrees
***                         6  -   Flip Image and Rotate 180 degrees
***                         7  -   Flip Image and Rotate 270 degrees
***    Return:
************************************************************************

DECLARE INTEGER RotateImage IN wwImaging.dll AS _RotateImage ;
   STRING lcSource, INTEGER FlipType   
RETURN IIF(;
 _RotateImage(STRCONV(FULLPATH(lcImage) +CHR(0) ,5),;
              @lnFlipType) = 1,.T.,.F.)
ENDFUNC
* GetImageInfo


************************************************************************
* wwAPI :: WriteImage
****************************************
***  Function: Writes one image into another
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WriteImage(lcSource, lcInsert, lnLeft, lnTop, llNonOpaque)

DECLARE INTEGER WriteImage IN wwImaging.dll AS _WriteImage ;
   STRING lcSource, string lcInsert, ;
   INTEGER lnLeft, INTEGER lnTop, INTEGER lnOpaque
   
RETURN IIF(;
 _WriteImage(STRCONV(FULLPATH(lcSource) +CHR(0),5),;
			    STRCONV(FULLPATH(lcInsert) +CHR(0),5),;
             lnLeft, lnTop,IIF(llNonOpaque,0,1)) = 1,.T.,.F.)

ENDFUNC
*  wwAPI :: WriteImage

************************************************************************
* wwAPI :: ReadImage
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ReadImage(lcSource, lcTarget, lnLeft, lnTop, lnWidth, lnHeight)

DECLARE INTEGER ReadImage IN wwImaging.dll AS _ReadImage ;
   STRING lcSource, STRING lcTarget, INTEGER lnLeft, INTEGER lnTop, ;
   INTEGER lnWidth, INTEGER lnHeight
   
RETURN IIF(;
 _ReadImage(STRCONV(FULLPATH(lcSource) +CHR(0),5),;
			 STRCONV(FULLPATH(lcTarget) +CHR(0),5),;
             lnLeft, lnTop, lnWidth, lnHeight) = 1,.T.,.F.)
ENDFUNC
*  wwAPI :: ReadImage

************************************************************************
* wwAPI ::  GetCaptchaImage
****************************************
***  Function: Returns an image for the given text
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetCaptchaImage(lcText,lcOutputFile,lcFont,lnFontSize)

IF EMPTY(lcFont)
   lcFont = "Arial"
ENDIF
IF EMPTY(lnFontSize)
   lnFontSize = 28
ENDIF      

DECLARE INTEGER GetCaptchaImage ;
   IN wwImaging.dll  as _GetCaptchaImage ;
   STRING Text,  STRING FONTNAME, integer FontSize, STRING lcOutputFile

lcText = STRCONV(lcText,5) + CHR(0)
lcFont = STRCONV(lcFont,5) + CHR(0)
lcOutputFile = STRCONV(lcOutputFile,5) + CHR(0)   
   
RETURN IIF( _GetCaptchaImage(lcText,lcFont,lnFontSize,lcOutputFile) = 1,;
           .T.,.F.)
ENDFUNC
*  wwAPI ::  GetCaptchaImage

************************************************************************
*  MapNetworkDrive
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION MapNetworkDrive(lcDrive, lcSharePath, lcUsername, lcPassword)

IF RIGHT(lcDrive,1) != ":"
   lcDrive = lcDrive + ":"
ENDIF
   
lcRun = [net use ] + lcDrive + [ "] + lcSharePath + [" ]

IF !EMPTY(lcUsername)
  lcUserName = ["]  + lcPassword + [" /USER:"] + lcUsername + ["]
ENDIF

lcRun = lcRun + lcUsername
RUN &lcRun 

ENDFUNC
*   MapNetworkDrive

************************************************************************
* wwAPI :: CreateprocessEx
****************************************
***  Function: Calls the CreateProcess API to run a Windows application
***    Assume: Gets around RUN limitations which has command line
***            length limits and problems with long filenames.
***            Can do Redirection
***            Requires wwIPStuff.dll to run!
***      Pass: lcExe - Name of the Exe
***            lcCommandLine - Any command line arguments
***    Return: .t. or .f.
************************************************************************
FUNCTION CreateProcessEx(lcExe,lcCommandLine,lcStartDirectory,;
                         lnShowWindow,llWaitForCompletion,lcStdOutputFilename)
LOCAL lnWait, lnResult

DECLARE INTEGER wwCreateProcess IN wwIPStuff.DLL AS _wwCreateProcess  ;
   String lcExe, String lcCommandLine, INTEGER lnShowWindow,;
   INTEGER llWaitForCompletion, STRING lcStartupDirectory, STRING StdOutFile
   
IF EMPTY(lcStdOutputFileName)
  lcStdOutputFileName = NULL
ENDIF
IF EMPTY(lcStartDirectory)
  lcStartDirectory = NULL
ENDIF

IF !EMPTY(lcCommandLine)
   lcCommandLine = ["] + lcExe + [" ]+ lcCommandLine
ELSE
   lcCommandLine = ""
ENDIF

IF llWaitForCompletion
   lnWait = 1
ELSE
   lnWait = 0
ENDIF
IF VARTYPE(lnShowWindow) # "N"
   lnShowWindow = 4
ENDIF   

lnResult = _wwCreateProcess(lcExe,lcCommandLine,lnShowWindow,lnWait,lcStartDirectory,lcStdOutputFileName)  

RETURN IIF(lnResult == 1, .t. , .f.)
ENDFUNC

************************************************************************
* wwAPI :: Createprocess
****************************************
***  Function: Calls the CreateProcess API to run a Windows application
***    Assume: Gets around RUN limitations which has command line
***            length limits and problems with long filenames.
***            Can do everything EXCEPT REDIRECTION TO FILE!
***      Pass: lcExe - Name of the Exe
***            lcCommandLine - Any command line arguments
***    Return: .t. or .f.
************************************************************************
FUNCTION Createprocess(lcExe,lcCommandLine,lnShowWindow,llWaitForCompletion)
LOCAL hProcess, cProcessInfo, cStartupInfo

DECLARE INTEGER CreateProcess IN kernel32 as _CreateProcess; 
    STRING   lpApplicationName,; 
    STRING   lpCommandLine,; 
    INTEGER  lpProcessAttributes,; 
    INTEGER  lpThreadAttributes,; 
    INTEGER  bInheritHandles,; 
    INTEGER  dwCreationFlags,; 
    INTEGER  lpEnvironment,; 
    STRING   lpCurrentDirectory,; 
    STRING   lpStartupInfo,; 
    STRING @ lpProcessInformation 

 
cProcessinfo = REPLICATE(CHR(0),128)
cStartupInfo = GetStartupInfo(lnShowWindow)

IF !EMPTY(lcCommandLine)
   lcCommandLine = ["] + lcExe + [" ]+ lcCommandLine
ELSE
   lcCommandLine = ""
ENDIF

lnResult =  _CreateProcess(lcExe,lcCommandLine,0,0,1,0,0,;
                           SYS(5)+CURDIR(),cStartupInfo,@cProcessInfo)

lhProcess = CHARTOBIN( SUBSTR(cProcessInfo,1,4) )

IF llWaitForCompletion
   #DEFINE WAIT_TIMEOUT 0x00000102
   DECLARE INTEGER WaitForSingleObject IN kernel32.DLL ;
         INTEGER hHandle, INTEGER dwMilliseconds

   DO WHILE .T.
       *** Update every 100 milliseconds
       IF WaitForSingleObject(lhProcess, 100) != WAIT_TIMEOUT
          EXIT
        ELSE
           DOEVENTS
        ENDIF
   ENDDO
ENDIF


DECLARE INTEGER CloseHandle IN kernel32.DLL ;
        INTEGER hObject

CloseHandle(lhProcess)

RETURN IIF(lnResult=1,.t.,.f.)


************************************************************************
*  GetStartupInfo
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetStartupInfo(lnShowWindow)
LOCAL lnFlags
* creates the STARTUP structure to specify main window
* properties if a new window is created for a new process

IF EMPTY(lnShowWindow)
  lnShowWindow = 1
ENDIF
  
*| typedef struct _STARTUPINFO {
*| DWORD cb; 4
*| LPTSTR lpReserved; 4
*| LPTSTR lpDesktop; 4
*| LPTSTR lpTitle; 4
*| DWORD dwX; 4
*| DWORD dwY; 4
*| DWORD dwXSize; 4
*| DWORD dwYSize; 4
*| DWORD dwXCountChars; 4
*| DWORD dwYCountChars; 4
*| DWORD dwFillAttribute; 4
*| DWORD dwFlags; 4
*| WORD wShowWindow; 2
*| WORD cbReserved2; 2
*| LPBYTE lpReserved2; 4
*| HANDLE hStdInput; 4
*| HANDLE hStdOutput; 4
*| HANDLE hStdError; 4
*| } STARTUPINFO, *LPSTARTUPINFO; total: 68 bytes

#DEFINE STARTF_USESTDHANDLES 0x0100
#DEFINE STARTF_USESHOWWINDOW 1
#DEFINE SW_HIDE 0
#DEFINE SW_SHOWMAXIMIZED 3
#DEFINE SW_SHOWNORMAL 1

lnFlags = STARTF_USESHOWWINDOW

RETURN binToChar(80) +;
    binToChar(0) + binToChar(0) + binToChar(0) +;
    binToChar(0) + binToChar(0) + binToChar(0) + binToChar(0) +;
    binToChar(0) + binToChar(0) + binToChar(0) +;
    binToChar(lnFlags) +;
    binToWordChar(lnShowWindow) +;
    binToWordChar(0) + binToChar(0) +;
    binToChar(0) + binToChar(0) + binToChar(0) + REPLICATE(CHR(0),30)


************************************************************************
*  GetDotNetFrameworkPath
****************************************
***  Function: Returns the highest active .NET version path
***            Note returns 2.0 version for 3.0 or 3.5 since
***            those versions are not full frameowrk versions.
***      Pass: nothing
***    Return: Path with trailing backslash or "" on failure
************************************************************************
FUNCTION GetDotNetFrameworkPath()
LOCAL lcString, lnSize, lnResult, llError

llError = .F.
TRY
	DECLARE INTEGER GetCORSystemDirectory IN MSCorEE.dll ;
	    string @, integer, integer@	    

	lcString = SPACE(512)
	lnSize = LEN(lcString)
	    
	lnResult = GetCorSystemDirectory(@lcString,lnSize,@lnSize)
	IF lnResult # 0
	   RETURN ""
	ENDIF   
CATCH
   
ENDTRY

IF llError
	RETURN ""
ENDIF

RETURN LEFT( STRCONV(lcString,6),lnSize-1)
ENDFUNC
*   GetDotNetFrameworkPath


************************************************************************
* Sleep
****************************************
***  Function: Suspends the current thread for x number of 
***            milliseconds.
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WinAPI_Sleep(lnMilliSecs)

lnMillisecs=IIF(type("lnMillisecs")="N",lnMillisecs,0)

DECLARE Sleep ;
  IN WIN32API ;
  INTEGER nMillisecs
    
=Sleep(lnMilliSecs)    
ENDFUNC
* WinApi_Sleep

************************************************************************
*  WinApi_GetWindowRect
****************************************
***  Function: Returns the size of a window returning top, left, bottom right
***            locations in pixel sizes
***      Pass: HWND for a window
***    Return: Object with Left,Top,Right,Bottom values
************************************************************************
FUNCTION WinApi_GetWindowRect(lnHwnd)

DECLARE INTEGER GetWindowRect IN user32;
    INTEGER hWindow,;
    STRING @lpRect
    
cBuffer = REPLICATE(CHR(0), 16)
GetWindowRect( lnHwnd, @cBuffer )

loRect = CREATEOBJECT("EMPTY")
ADDPROPERTY(loRect,"Left",CharToBin( SUBSTR(cBuffer, 1, 4) ))
ADDPROPERTY(loRect,"Top", CharToBin( SUBSTR(cBuffer, 5, 4) ))
ADDPROPERTY(loRect,"Right",CharToBin( SUBSTR(cBuffer, 9, 4) ))
ADDPROPERTY(loRect,"Bottom", CharToBin( SUBSTR(cBuffer, 13, 4) ))

RETURN loRect
ENDFUNC
*   WinApi_GetWindowRect

************************************************************************
*  WindowsVersion
****************************************
***  Function: Returns the Windows Version number as a number (6.1 for Win7)
***    Assume: http://www.nirmaltv.com/2009/08/17/windows-os-version-numbers/
***      Pass: nothing
***    Return: Version number as a number
************************************************************************
FUNCTION WindowsVersion()
RETURN VAL(OS(3) + "." + OS(4))
ENDFUNC
*   WindowsVersion

************************************************************************
*  WinApi_NullString
****************************************
***  Function: Strips a string to the first null value
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WinAPI_NullString(lcInput)

lnAt = AT(CHR(0),lcInput)
IF lnAt > 0
   lcInput = SUBSTR(lcInput,1,lnAt-1)
   RETURN lcInput
ENDIF

RETURN lcInput   
ENDFUNC
*   Win32_StripNull

************************************************************************
* wwAPI :: Win32_GetSystemTime
****************************************
***  Function: Returns the System local time (in UTC format)
***            for the current time
***    Assume:
***      Pass: 
***    Return:
************************************************************************
FUNCTION Win32_GetSystemTime()
LOCAL lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond, lcBuffer

DECLARE INTEGER GetSystemTime IN win32api STRING @
lcBuffer=SPACE(40)
=GetSystemTime(@lcBuffer)

#IF wwVFPVersion > 8
   lnYear = CTOBIN( SUBSTR(lcBuffer,1,2),"RS2")
   lnMonth = CTOBIN( SUBSTR(lcBuffer,3,2),"RS2")
   lnDay = CTOBIN( SUBSTR(lcBuffer,7,2),"RS2")

   lnHour = CTOBIN( SUBSTR(lcBuffer,9,2),"RS2")
   lnMinute = CTOBIN( SUBSTR(lcBuffer,11,2),"RS2")
   lnSecond = CTOBIN( SUBSTR(lcBuffer,13,2),"RS2")
#ELSE
   lnYear = CharToBin( SUBSTR(lcBuffer,1,2))
   lnMonth = CharToBin( SUBSTR(lcBuffer,3,2))
   lnDay = CharToBin( SUBSTR(lcBuffer,7,2))

   lnHour = CharToBin( SUBSTR(lcBuffer,9,2))
   lnMinute = CharToBin( SUBSTR(lcBuffer,11,2))
   lnSecond = CharToBin( SUBSTR(lcBuffer,13,2))
#ENDIF

lcTime = "{^" + TRANSFORM(lnYear) + "-" + ;
         TRANSFORM(lnMonth) + "-" + ;
         TRANSFORM(lnDay) + " " +;
         TRANSFORM(lnHour) + ":" +;
         TRANSFORM(lnMinute) + ":" + ;
         TRANSFORM(lnSecond) + "}"

RETURN  EVALUATE(lcTime)

************************************************************************
* wwAPI :: WinApi_ActivateWindow
****************************************
***  Function: Activates the 
***    Assume:
***      Pass: lcTitle - Exact Window Title or Window Handle Number
***    Return:
************************************************************************
FUNCTION ActivateWindow(lcTitle,lnParentHandle)

IF VARTYPE(lcTitle) = "C"
   IF EMPTY(lnParentHandle)
      lnParentHandle = 0
   ENDIF
   
   DECLARE INTEGER FindWindow ;
      IN WIN32API ;
      STRING cNull,STRING cWinName

   lnHandle = FindWindow(lnParentHandle,lcTitle)
ELSE
   lnHandle = lcTitle
ENDIF

DECLARE INTEGER SetForegroundWindow ;
      IN WIN32API INTEGER

SetForegroundWindow(lnHandle)

RETURN
ENDFUNC
*  wwAPI :: WinApi_ActivateWindow


#DEFINE SHACF_AUTOAPPEND_FORCE_OFF  0x80000000
#DEFINE SHACF_AUTOAPPEND_FORCE_ON  0x40000000
#DEFINE SHACF_AUTOSUGGEST_FORCE_OFF  0x20000000
#DEFINE SHACF_AUTOSUGGEST_FORCE_ON  0x10000000
#DEFINE SHACF_DEFAULT  0x0
#DEFINE SHACF_FILESYSTEM  0x1
#DEFINE SHACF_URLHISTORY  0x2
#DEFINE SHACF_URLMRU  0x4
#DEFINE SHACF_URLALL  SHACF_URLHISTORY +SHACF_URLMRU


************************************************************************
*  ActivateFileSystemAutoCompletion
****************************************
***  Function: Activates AutoComplete on an edit control
***    Assume:
***      Pass: loEditControl -  Any input control
***    Return:
************************************************************************
FUNCTION ActivateFileSystemAutoCompletion(loEditControl)

DECLARE SHAutoComplete IN "shlwapi.dll" LONG hwndEdit, LONG dwFlags
SHAutoComplete(loEditControl.HWND , SHACF_AUTOSUGGEST_FORCE_ON +SHACF_FILESYSTEM)

ENDFUNC
*   WinApi_ActivateDirectoryAutoComplete

************************************************************************
*  ActivateUrlAutoCompletion
****************************************
***  Function: Activates AutoComplete on an edit control
***    Assume:
***      Pass: loEditControl -  Any input control
***    Return:
************************************************************************
FUNCTION ActivateUrlAutoCompletion(loEditControl)

DECLARE SHAutoComplete IN "shlwapi.dll" LONG hwndEdit, LONG dwFlags
SHAutoComplete(loEditControl.HWND , SHACF_AUTOSUGGEST_FORCE_ON +SHACF_URLHISTORY)

ENDFUNC
*   WinApi_ActivateDirectoryAutoComplete




************************************************************************
* InstallPrinterDriver
*********************************************
***  Function: Installs a Windows Printer driver from the stock
*** 		   Windows driver library
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION InstallPrinterDriver(lcDriverName,lcPrinterName)
LOCAL lcOs, llResult

IF EMPTY(lcDriverName)
  *** This color PS driver exists under Win2003, Vista and XP
  lcDriverName = "Xerox Phaser 1235 PS"
  
  * lcDriverName = "Apple Color LW 12/660 PS"
  DO CASE 
	  *** Windows 7  (V 6.1)
	  CASE OS(3) = "6" AND OS(4) = "1"
      	lcDriverName = "Xerox Phaser 6120 PS" 
	  CASE OS(3) = "6" AND OS(4) > "1"
      	lcDriverName = "Xerox PS Color Class Driver" 
  ENDCASE
ENDIF

IF EMPTY(lcPrinterName)
  lcPrinterName = lcDriverName
ENDIF  
LOCAL ARRAY laPrinters[1]

lnCount = APRINTERS(laPrinters)
FOR lnX = 1 TO lnCount
   IF LOWER(lcPrinterName) == LOWER(laPrinters[lnX,1])
      RETURN .t.
   ENDIF
ENDFOR   

lcOS = "Windows 2000 or XP"
IF OS(3) = "6"  && Vista requires XP setting
   lcOS = "Windows XP"
ENDIF
   

loAPI = CREATEOBJECT("wwAPI")
lcExe = loAPI.GetSystemDir() + "rundll32.exe"
lcCmdLine = [printui.dll,PrintUIEntry /ia /m "] + lcDriverName + [" /h "Intel" /v "] + lcOS + [" /f "] + loAPI.GetSystemDir(.t.)+ [inf\ntprint.inf" /u"]

llResult = CreateProcess(lcExe,lcCmdLine,2,.t.)
IF !llResult
   RETURN .F.
ENDIF   

lcCmdLine = [printui.dll,PrintUIEntry /if /b "] + lcPrinterName + [" /f "] + loAPI.GetSystemDir(.T.) + [inf\ntprint.inf" /r "lpt1:" /m "] + lcDriverName + ["]

llResult = CreateProcess(lcExe,lcCmdLine,2,.t.)
IF !llResult
   RETURN .F.
ENDIF   

RETURN .T.

************************************************************************
*  FindPostScriptPrinter
****************************************
***  Function: Tries to find a PostScriptPrinter for using with 
***            PDF generation solutions like GhostScript
***            Looks for a few known PS drivers and uses that.
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION FindPostScriptPrinter(lcDefault)
LOCAL ARRAY laPrinters[1,2]

IF EMPTY(lcDefault)
  lcDefault = "Xerox Phaser 1235 PS"
ENDIF
  
lnPrinters = APRINTERS(laPrinters)
FOR lnX = 1 TO lnPrinters
  lcPrinter = laPrinters[lnX]
  DO CASE
  	 CASE lcPrinter = "Apple" AND  RIGHT(lcPrinter,3) = " PS"
  	 	RETURN lcPrinter
  	 CASE lcPrinter = "Xerox" AND ATC(" PS",lcPrinter) > 0
  	    RETURN lcPrinter  	    
  ENDCASE
ENDFOR

*** Try any printer that ends in PS
FOR lnX = 1 TO lnPrinters
  lcPrinter = laPrinters[lnX]
  IF RIGHT(lcPrinter,3) = " PS"
     RETURN lcPrinter
  ENDIF
ENDFOR

RETURN lcDefault
ENDFUNC
*   FindPostScriptPrinter


************************************************************************
* wwAPI ::  GetClipboardText
****************************************
***  Function:
***    Assume:
***      Pass: lcFormat -  "Rich Text Format", "Html Format"
***    Return: Specified text or NULL
************************************************************************
FUNCTION GetClipboardText(lcFormat)
LOCAL lcDoc, lnSize

DECLARE INTEGER GetClipboardText IN wwipstuff.dll  as _GetClipboardText;
    String@, INTEGER@, String

lcDoc = SPACE(65356)
lnSize = LEN(lcDoc)
IF ( _GetClipboardText(@lcDoc,@lnSize,lcFormat) # 1 )
   IF lnSize > 0
	  *** Retry with a bigger buffer
      lcDoc = SPACE(lnsize)
      IF _GetClipboardText(@lcDoc,@lnSize,lcFormat) # 1
         RETURN NULL
      ENDIF
   ELSE
     RETURN null
   ENDIF
ENDIF
RETURN LEFT(lcDoc,lnSize)
ENDFUNC
*  wwAPI ::  GetClipboardText


************************************************************************
* wwAPI :: WinApi_SendMessage
****************************************
***  Function: SendMessage API call - straight through
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WinApi_SendMessage(lnHwnd,lnMsg,lnWParam,lnLParam)

DECLARE integer SendMessage IN WIN32API ;
        integer hWnd,integer Msg,;
        integer wParam,;
        Integer lParam

RETURN SendMessage(lnHwnd,lnMsg,lnWParam,lnLParam)
ENDFUNC  

************************************************************************
* wwAPI :: WinApi_FindWindowEx
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WinApi_FindWindowEx(lnParentHwnd,lnHwndLastChild,lcClass,lcTitle)

  IF EMPTY(lcClass)
     lcClass = NULL
  ENDIF
  IF EMPTY(lcTitle)
     lcTitle = NULL
  ENDIF
  IF EMPTY(lnHwndLastChild)
    lnHwndLastChild = 0
  ENDIF
  
  declare integer FindWindowEx in Win32API;
       integer, integer, string, string

RETURN FindWindowEx(lnParentHwnd,lnHwndLastChild,lcClass,lcTitle)
ENDFUNC

************************************************************************
*  WinApi_MoveWindow
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WinApi_MoveWindow(lnHandle, lnX, lnY, lnWidth, lnHeight)

DECLARE INTEGER MoveWindow IN WIN32API ;
	INTEGER hwnd,;
	INTEGER x,;
	INTEGER y,;
	INTEGER width, ;
	Integer height, ;
	Integer repaint

RETURN MoveWindow(lnHandle, lnX,lnY,lnWidth,lnHeight)
ENDFUNC
*   WinApi_MoveWindow

************************************************************************
* wwAPI :: WinApi_GetClassName
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WinApi_GetClassName(lnHwnd)

 declare integer GetClassName in Win32API ;
  integer lnhWnd, string @lpClassName, integer lnMaxCount

   lnBuffer   = 255
   lcBuffer   = space(lnBuffer)
   lnBuffer   = GetClassName(lnhWnd, @lcBuffer, lnBuffer)
   IF lnBuffer > 0
      RETURN left(lcBuffer, lnBuffer - 1)
   ENDIF
   
   RETURN ""
ENDFUNC

************************************************************************
* wwAPI :: WinApi_CallWindowProc
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WinApi_CallWindowProc(lpLastWinProc,lnHwnd,lnMsg,lnWParam,lnLParam)
declare integer CallWindowProc in Win32API ;
         integer lpPrevWndFunc, integer hWnd, integer Msg,;
         integer wParam, integer lParam

RETURN CallWindowProc(lhLastWinProc,lnHwnd,lnMsg,lnWParam,lnLParam)
ENDFUNC    

************************************************************************
* wwAPI :: WinApi_GetWindowLong
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WinApi_GetWindowLong(lnHwnd,lnIndex)
   declare integer GetWindowLong in Win32API ;
         integer hWnd, integer nIndex

   IF VARTYPE(lnIndex) # "N"
      lnIndex = -4  &&GWL_WNDPROC  
   ENDIF

RETURN GetWindowLong(lnHwnd,lnIndex)        
ENDFUNC
*  wwAPI :: WinApi_GetWindowLong     

************************************************************************
* wwAPI :: Sleep
*********************************
***  Function: Puts the computer into idle state. More efficient and
***            no keyboard interface than Inkey()
***      Pass: tnMilliseconds
***    Return: nothing
************************************************************************
FUNCTION Sleep(lnMilliSecs)

lnMillisecs=IIF(type("lnMillisecs")="N",lnMillisecs,0)

DECLARE Sleep ;
  IN WIN32API ;
  INTEGER nMillisecs
 	
=Sleep(lnMilliSecs) 	
ENDFUNC

************************************************************************
*  Is64Bit
****************************************
***  Function: Determines whether you're running on 64 bit Windows
***    Assume:
***      Pass:
***    Return: .T. or .F.
************************************************************************
FUNCTION Is64Bit()

loAPI = CREATEOBJECT("wwAPI")
lcWin = loAPI.GetSystemDir(.T.)

IF  IsDir(lcWin + "SysWOW64")
   RETURN .T.
ENDIF

RETURN .F.
* Is64Bit


************************************************************************
* GetMonitorStatistics
****************************************
***  Function: Returns information about the desktop screen
***            Can be used to check for desktop width and size
***            and determine when a second monitor is disabled
***            and layout needs to be adjusted to keep the
***            window visible.
***      Pass:
***    Return:  Monitor Object
************************************************************************
FUNCTION GetMonitorStatistics()

#DEFINE SM_XVIRTUALSCREEN 76
#DEFINE SM_YVIRTUALSCREEN 77
#DEFINE SM_CXVIRTUALSCREEN 78
#DEFINE SM_CYVIRTUALSCREEN 79
#DEFINE SM_CMONITORS 80
#DEFINE SM_CXFULLSCREEN 16
#DEFINE SM_CYFULLSCREEN 17


DECLARE INTEGER GetSystemMetrics IN user32 INTEGER nIndex

loMonitor = CREATEOBJECT("EMPTY")
ADDProperty( loMonitor,"Monitors",GetSystemMetrics(SM_CMONITORS) )
ADDPROPERTY( loMonitor,"VirtualWidth",GetSystemMetrics(SM_CXVIRTUALSCREEN) )
ADDPROPERTY( loMonitor,"VirtualHeight",GetSystemMetrics(SM_CYVIRTUALSCREEN) )
ADDPROPERTY( loMonitor,"ScreenHeight",GetSystemMetrics(SM_CYFULLSCREEN) )
ADDPROPERTY( loMonitor,"ScreenWidth",GetSystemMetrics(SM_CXFULLSCREEN) )

RETURN loMonitor
ENDFUNC
*  wwAPI ::  GetMonitorStatistics

************************************************************************
*  FixMonitorPosition
****************************************
***  Function: Fixes a FoxPro form to fit on the screen and become
***            visible on activation even if the location is on
***            no longer visible screen
***            This function is useful if you store screen positions
***            in configuration files and have a screen on a second
***            monitor that is no longer available
***    Assume:
***      Pass: loForm    -  The form to fix
***    Return: nothing
************************************************************************
FUNCTION FixMonitorPosition(loForm,lnWidth, lnHeight)
LOCAL loMonitor

*** Retrieve statistics about virtual and active screen
loMonitor = GetMonitorStatistics()

IF EMPTY(lnWidth)
   lnWidth = loMonitor.VirtualWidth - 10
ENDIF 
IF EMPTY(lnHeight)
   lnHeight = loMonitor.VirtualHeight - 10
ENDIF   

*** If the monitor is on a non-visible screen move it over
*** to the current screen

*** Fix top and left first - on another screen most likely
IF loForm.Left > loMonitor.VirtualWidth - 10
	loForm.Left = 5
ENDIF
IF loForm.Top > loMonitor.VirtualHeight - 50
	loForm.Top = 5
ENDIF

*** Now fix the width if larger than screen
IF loForm.Width > loMonitor.VirtualWidth - 10
   loForm.Width = lnWidth -  10
   loForm.Left = 5
ENDIF
IF loForm.Height > loMonitor.VirtualHeight - 10
   loForm.Height = lnHeight - 10
   loForm.Top = 5
ENDIF

ENDFUNC
*   FixMonitorPosition


************************************************************************
* wwUtils :: UnZipFiles
*********************************
***  Function: Unzips files to a specified directory
***    Assume: Requires DynaZip DLLs (dunzip32.dll)
***      Pass: lcZipFile
***            lcDestination  -  Dir to unzip to
***            lcFileSpec     -  Files to unzip (*.*)
***    Return: DynaZip Error Code or 0 on success
************************************************************************
FUNCTION UnZipFiles
LPARAMETERS lcZipFile, lcDestination, lcFileSpec

lcFileSpec=IIF(type("lcFileSpec")="C",lcFileSpec,"*.*")
lcDestination=IIF(type("lcDestination")="C",lcDestination,SYS(5) + CURDIR())

DECLARE INTEGER UnZip ;
   IN wwipstuff.dll ;
   STRING ZipFile,;
   STRING Destination,;
   STRING FileSpec

RETURN UnZip(lcZipFile,lcDestination,lcFileSpec)

************************************************************************
* wwUtils :: ZipFiles
*********************************
***  Function: Zips files
***    Assume: Function requires DynaZip DLLs (dzip32.dll)
***      Pass: lcZipFile   - Fully qualified ZIP file name 
***            lcFileList  - Comma Delimited file list (Wildcards OK)
***    Return: DynaZip error code or 0
************************************************************************
FUNCTION ZipFiles
LPARAMETERS lcZipFile, lcFileList, lnCompression, llRecurse, llAdditive

lnCompression=IIF(type("lnCompression")="N",lnCompression,9)
 
DECLARE INTEGER Zip ;
   IN wwipstuff.dll ;
   STRING ZipFile,;
   STRING FileList,;
   INTEGER lnCompression,;
   INTEGER lnRecurse,;
   INTEGER lnAdditive
   

RETURN Zip(lcZipFile,lcFileList,lnCompression,IIF(llRecurse,1,0),IIF(llAdditive,1,0) )

************************************************************************
* wwApi :: DecodeDBF
*********************************
***  Function: Decodes a DBF file encoded EncodeDBF back into its
***            DBF/FPT format
***      Pass:
***    Return:
************************************************************************
FUNCTION DecodeDBF
LPARAMETERS lcBuffer,lcDBF
LOCAL lnSeparator, lcHeader, lcFname, lnSize1, lnSize2, lcDBF, lcFile1, lcFile2

   IF LEN(lcBuffer)<105
      RETURN .F.
   ENDIF

   lcHeader=SUBSTR(lcBuffer,1,105)
   lcFname=TRIM(SUBSTR(lcBuffer,6,40))
   lnSize1=VAL(SUBSTR(lcBuffer,46,10))
   lnSize2=VAL(SUBSTR(lcBuffer,96,10))

   *** Use parm or the filename specified in the header
   lcDBF=IIF(EMPTY(lcDBF),lcFname,UPPER(lcDBF))

   IF lcHeader # "wwDBF"
      RETURN .F.
   ENDIF

   lcFile1=""
   lcFile2=""

   IF lnSize1 > 0
      lcFile1=SUBSTR(lcBuffer,106,lnSize1)
      IF LEN(lcFile1) < lnSize1
         RETURN .F.
      ENDIF
   ENDIF
   IF lnSize2 > 0
      lcFile2=SUBSTR(lcBuffer,106 + lnSize1, lnSize2)
      lnSizex=LEN(lcFile2)
      IF LEN(lcFile2) < lnSize2 - 1
         RETURN .F.
      ENDIF
   ENDIF

   =File2Var(lcDBF,lcFile1)

   IF !EMPTY(lcFile2)
      =File2Var(STRTRAN(lcDBF,".DBF",".FPT"),lcFile2)
   ENDIF

RETURN .T.
ENDFUNC
* wwHTTP :: DecodeDBF


********************************************************
* wwAPI :: EncodeDBF
*********************************
***  Function: This function encodes a DBF file ready to
***            be sent up to a server using HTTPGetEx in
***            the POST buffer. The file will be URL
***            encoded.
***    Assume: Note you can send a ZIP file here, too!
***            105 byte header on top of file contains
***            5 byte ID (wwDBF) filename (40 bytes) and
***            size(10 bytes) for each file
***      Pass: lcDBF     - Full DBF filename w/ ext
***            llHasMemo - .t. or (.f.)
***    Return: Encoded Buffer or "" on failure
********************************************************
FUNCTION EncodeDBF
LPARAMETERS lcDBF, llHasMemo, lcEncodedName
LOCAL lcBuffer1, lcBuffer2, lcDBF, lcHeader, lcFPT

lcDBF=IIF(VARTYPE(lcDBF)="C",UPPER(lcDBF),"")
IF EMPTY(lcEncodedName)
   lcEncodedName = JUSTFNAME(lcDBF) 
ENDIF

IF !FILE(lcDBF)
   RETURN ""
ENDIF

lcBuffer1=File2Var(lcDBF)
lcHeader = "wwDBF" + PADR(lcEncodedName,40) + ;
   STR(LEN(lcBuffer1),10)
IF !llHasMemo
   lcHeader=lcHeader+ SPACE(50)  && Pad out header
   RETURN lcHeader + lcBuffer1
ENDIF

lcFPT=STRTRAN(LOWER(lcDBF),".dbf",".fpt")

lcBuffer2=File2Var(lcFPT)
lcHeader=lcHeader + PADR(FORCEEXT(lcEncodedName,"fpt"),40) + ;
   STR(LEN(lcBuffer2),10)

RETURN lcHeader + lcBuffer1 + lcBuffer2
ENDFUNC

FUNCTION EncodeFile(lcFile)
   RETURN EncodeDbf(lcFile)
RETURN
FUNCTION DecodeFile(lcBuffer)
  RETURN DecodeDbf(@lcBuffer)
ENDFUNC  


*** Internet Dialing and PING routines


************************************************************************
* GetDomainFromIp
*********************************
***  Function: Returns the domain name from an IP Addressed passed
***    Assume: Can be slow as a reverse look up against DNS server
***            is made. Over dialup this can take 5-10 seconds for
***            initial connection about 1 second once connected.
***      Pass: lcIPAddress   -   An IP Address (111.111.111.111)
***    Return: Domain Name or "" if !found or failure
************************************************************************
FUNCTION GetDomainFromIp
LPARAMETERS lcIPAddress
LOCAL lcDomain

DECLARE GetDomainFromIp  ;
     IN wwIPstuff.dll as GetDomainFromIp_API ;
     String @cDomain,;
     String cIpAddress

lcDomain=SPACE(200)   
lnValue= GetDomainFromIp_API(@lcDomain,lcIPAddress) 
IF AT(CHR(0),lcDomain) > 1
  lcDomain=substr(lcDomain,1,AT(chr(0),lcDomain) -1 )
ELSE
  lcDomain=""
ENDIF
  
RETURN lcDomain  
ENDFUNC
*   GetDomainFromIp


************************************************************************
* GetIpFromDomain
*********************************
***  Function: Returns the IP address of a given domain name.
***      Pass: lcDomain  -    Domain Name address
***    Return: IP Address of "" if not resolved
************************************************************************
FUNCTION GetIpFromDomain
LPARAMETERS lcDomain

DECLARE GetIpFromDomain ;
  IN wwipstuff.dll  AS GetIpFromDomain_API ;
  String @cIpAddress,;
  String cDomain

lcIpAddress=SPACE(56)

lnValue= GetIpFromDomain_API(@lcIpAddress,lcDomain) 
IF AT(CHR(0),lcIpAddress) > 1
  lcIpAddress=substr(lcIpAddress,1,AT(chr(0),lcIpADdress) -1 )
ELSE
  lcIpAddress=""
ENDIF

RETURN lcIpAddress
ENDFUNC


PROCEDURE rasdial
LPARAMETERS lcRASEntry, lcUsername, lcPassword
LOCAL lcErrorMsg, lnResult

DECLARE INTEGER DoRasDial ;
   IN wwIPstuff.dll ;
   STRING szEntry,;
   STRING szUsername,;
   STRING szPassword,;
   INTEGER @dwHandle,;
   STRING @Errormsg

THIS.cErrorMsg = ""

lcErrorMsg = SPACE(256)

lnHandle=0
lnResult = DoRasDial(lcRASEntry,lcUserName,lcPassword,;
                     @lnHandle,@lcErrorMsg)

IF lnResult # 0
   THIS.nRASHandle = lnHandle
ELSE
   THIS.nRASHandle = 0   
   THIS.cErrorMsg = TRIM(STRTRAN(lcErrorMsg,CHR(0),""))
ENDIF   

RETURN IIF(lnResult = 0,.F.,.T.)
************************************************************************
* wwIPStuff :: RASDial
*********************************
***  Function: Connects via the RAS dialer - no interface except a 
***            MessageBox on error!
***     Notes: Connection handle is stored in THIS.nRASHandle. This
***            handle is necessary to close down a RAS Connection
***            via RASHangup.
***      Pass: lcRASEntry      -   Name of the RAS Connection to use
***            lcUsername	   -   Username to connect
***            lcPassword      -   Password to connect
***    Return: .t. or .f.
************************************************************************
ENDPROC


*-- Hangs up a previously dialed RAS connection. Requires handle!
PROCEDURE rashangup

DECLARE INTEGER DoRasHangup ;
   IN wwIPstuff.dll ;
   INTEGER lnHandle

IF THIS.nRASHandle = 0
   THIS.cErrorMsg = "No Connection active"
ENDIF
   
llResult = IIF( DoRasHangup(THIS.nRASHandle) = 0,.T.,.F.)
IF llResult
   THIS.nRASHandle = 0
ELSE
   THIS.cErrorMsg = "Invalid Connection Handle..."
ENDIF
  
RETURN llResult
************************************************************************
* wwIPStuff :: RASHangup
*********************************
***  Function: Terminates the RAS Connection that was initiated with
***            RASDial.
***    Assume: Works only on connections opened with RASDial and the
***            THIS.nRASHandle set there. Use a persistent object to
***            keep this reference around!
***    Return: .t. or .f.
************************************************************************
ENDPROC


*-- Dials an internet connection
PROCEDURE internetdial
	LPARAMETERS llUnattended, lcConnection

	IF EMPTY(lcConnection)
	   lcConnection = ""
	ENDIF

	IF llUnattended
	   lnDialMode = 0x8000 && INTERNET_DIAL_UNATTENDED
	ELSE
	   lnDialMode = 0
	ENDIF   

	DECLARE INTEGER InternetDial in WININET.DLL ;
	   INTEGER hwndParent,;
	   STRING lpszConnectoid,;
	   INTEGER dwFlag,;
	   INTEGER @lpdwConnection,;
	   INTEGER dwReserved
	   
	lnConn = 0
	lnResult = InternetDial(0,lcConnection,lnDialMode,@lnConn,0)

	IF lnResult = 0
	   THIS.nRasHandle = lnConn
	   RETURN .T.
	ENDIF

	THIS.nRasHandle = 0
	THIS.nerror=lnResult
	DO CASE
	 CASE lnResult = 631
	    this.cErrorMsg = "User Cancelled"
	 CASE lnResult = 668
	    this.cErrorMsg = "No Connection"
	 CASE lnResult = 87
	    this.cErrorMsg = "Invalid Parameters"
	 CASE lnResult = 680
	    this.cErrorMsg = "No Dialtone"
	 CASE lnResult = 678
	    this.cErrorMsg = "No answer"
	 CASE lnResult = 620
	    this.cErrorMsg = "No Endpoints - connection is misconfigured"
	 CASE lnResult = 623
	    this.cErrorMsg = "Couldn't find phonebook entry"
	 OTHERWISE
	   THIS.cerrormsg = THIS.GetSystemerrormsg()

	ENDCASE

	RETURN .F.
ENDPROC


*-- Hangs up Internet Connection
PROCEDURE internethangup
	LPARAMETERS lnHandle

	IF EMPTY(lnHandle)
	  lnHandle = this.nRasHandle
	ENDIF

	DECLARE INTEGER InternetHangUp IN WinInet.dll ;
	   INTEGER dwConnection, integer dwReserved
	   
	IF InternetHangUp(lnHandle,0) = 0
	   RETURN .T.
	ENDIF

	RETURN .F.   
ENDPROC

*-- 
************************************************************************
*  Ping
****************************************
***  Function: Pings a site and return hops and time optionally.
***    Assume:
***      Pass: lcDomain -  Domain or IP Address
***            @lnHops   -  Hops to the target site
***            @lnTime   -  Time for the ping in ms
***    Return: .T. or .f.
************************************************************************
FUNCTION Ping
LPARAMETERS lcDomain,lnHops, lnTime

lcIP = GetIpFromDomain(lcDomain)

DECLARE INTEGER GetRTTAndHopCount IN Iphlpapi;
		INTEGER DestIpAddress, LONG @HopCount,;
		INTEGER MaxHops, LONG @RTT

DECLARE INTEGER inet_addr IN ws2_32 STRING cp

LOCAL nDst, nHop, nRTT
nDst = inet_addr(lcIP) 
STORE 0 TO lnHops, lnTime

lnResult = GetRTTAndHopCount(nDst, @lnHops, 100, @lnTime)
IF lnResult <>0
   RETURN .T.
ENDIF

RETURN .F.
ENDFUNC

************************************************************************
* GZipCompressString
****************************************
***  Function: Compresses a string using GZip
***    Assume: Requires ZLIB1.DLL 
***      Pass:
***    Return:
************************************************************************
FUNCTION GZipCompressString(lcString,lnCompressionLevel)
LOCAL lcOutput, lcOutFile,lcInFile, lnHandle

*** 1 - 9
IF EMPTY(lnCompressionLevel)
   lnCompressionLevel = -1  && Default
ENDIF

*** Must write to files
lcOutFile = ADDBS(SYS(2023)) + SYS(2015) + TRANS(_vfp.ProcessId) + ".gz"
lcInFile = lcOutFile + ".in"

*** Failure to write the file
IF !FILE2VAR(lcInFile,lcString)
   RETURN ""
ENDIF

IF !VARTYPE(_GZipLoaded) = "L"
	GzipLibraries()
ENDIF

TRY
   lnHandle = gzopen(lcOutFile,"wb")
   IF (lnHandle < 0)
      RETURN ""
   ENDIF

   *** Set the compression level
   gzsetparams(lnHandle,lnCompressionLevel,0)

   gzwrite(lnHandle,lcString,LEN(lcString))
   gzclose(lnHandle)
CATCH
   IF lnHandle > -1
      gzclose(lnHandle)
   ENDIF
ENDTRY

lcOutput = FILETOSTR(lcOutFile)

ERASE (lcOutFile)
ERASE (lcInFile)

RETURN lcOutput


************************************************************************
* GZipUncompressString
****************************************
***  Function: Uncompresses a GZip string
***    Assume: 
***      Pass: lcCompressed   -  Compressed String
***            llIsFile       -  if .T. lcCompressed is a file
***    Return: decompressed string
************************************************************************
FUNCTION GZipUncompressString(lcCompressed,llIsFile)
LOCAL lcInFile, lcOutput, lnHandle

IF llIsFile
   *** Use parameter file name as input
   lcInFile = lcCompressed
ELSE
   lcInFile = ADDBS(SYS(2023)) + SYS(2015) + TRANSFORM(_VFP.ProcessId) + ".gz"
   *** Copy file to disk and use as input
   FILE2VAR(lcInFile,lcCompressed)
ENDIF

IF !VARTYPE(_GZipLoaded) = "L"
	GzipLibraries()
ENDIF

lcOutput = ""
TRY
   lnHandle = gzopen(lcInFile,"rb")
   IF (lnHandle < 1)
      RETURN ""
   ENDIF

   lcOutput = ""
   DO WHILE .T.
      lcBuffer = SPACE(65535)
      lnResult = gzread(lnHandle,@lcBuffer,LEN(lcBuffer))
      IF lnResult < 1
         EXIT
      ENDIF
      lcOutput = lcOutput + LEFT(lcBuffer,lnResult)
   ENDDO
CATCH
   * Nothing
FINALLY
   IF lnHandle > 0	
	   gzclose(lnHandle)
   ENDIF
   ERASE(lcInFile)
ENDTRY

RETURN lcOutput
* eof GZipUncompressString

************************************************************************
* GzipLibraries
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GzipLibraries()

PUBLIC _GZipLoaded
_GZipLoaded=.T.

* Opens file for writing
DECLARE LONG gzopen IN zlib1.dll ;
   STRING @ zFile ,;
   STRING @ zMode

* Writes data from a compressed file - gzip
DECLARE LONG gzwrite IN zlib1.dll ;
   LONG FILE ,;
   STRING @ uncompr,;
   LONG uncomprLen

*** Set options on the compression
DECLARE LONG gzsetparams IN zlib1.DLL ;
   LONG  gzFile,;
   INTEGER LEVEL,;
   INTEGER strategy

DECLARE LONG gzread IN zlib1.dll  ;
   LONG gzFile,;
   STRING @ buf,;
   LONG LEN

* Closes the file
DECLARE LONG gzclose IN zlib1.dll ;
   LONG FILE

RETURN


*************************************************************
DEFINE CLASS wwFileStream AS Custom
*************************************************************
*: Author: Rick Strahl
*:         (c) West Wind Technologies, 2012
*:Contact: http://www.west-wind.com
*:Created: 01/04/2012
*************************************************************
#IF .F.
*:Help Documentation
*:Topic:
Class wwFileStream

*:Description:

*:Example:
DO wwutils

loStream = CREATEOBJECT("wwFileStream")
*loStream = CREATEOBJECT("wwMemoryStream")
loStream.WriteLine("Hello World")
loStream.WriteLine("Off we go")
loStream.WriteFile("c:\photos\sailbig.jpg")

ShowText( loStream.ToString()) 
*:Remarks:

*:SeeAlso:


*:ENDHELP
#ENDIF

nHandle = 0
cFileName = "" 
nLength = 0


************************************************************************
*  Init
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Init()

this.cFileName = SYS(2023)  + "\" +  SYS(2015) + ".txt"
this.nHandle = FCREATE(this.cFileName)
this.nLength = 0

ENDFUNC
*   Init

************************************************************************
*  Destroy
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Destroy()
this.Dispose()
ENDFUNC
*   Destroy

************************************************************************
*  Dispose
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Dispose()

IF THIS.nHandle > 0
   TRY
   FCLOSE(this.nHandle)
   DELETE FILE (this.cFileName)
   CATCH
   ENDTRY
ENDIF
this.nLength = 0
ENDFUNC
*   Destroy

************************************************************************
*  Write
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Write(lcContent)
THIS.nLength = THIS.nLength + LEN(lcContent)
FWRITE(this.nHandle,lcContent)
ENDFUNC
*   Write

************************************************************************
*  WriteLine
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WriteLine(lcContent)
this.Write(lcContent)
this.Write(CHR(13) + CHR(10))
ENDFUNC
*   WriteLine

************************************************************************
*  WriteFile
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WriteFile(lcFileName)
lcFileName = FULLPATH(lcFileName)
this.Write(FILETOSTR( lcFileName ))
ENDFUNC
*   WriteFile

************************************************************************
*  ToString()
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ToString()
LOCAL lcOutput

FCLOSE(this.nHandle)
lcOutput = FILETOSTR(this.cFileName)

*** Reopen the file
this.nHandle = FOPEN(this.cFileName,1)
FSEEK(this.nHandle,0,2)

RETURN lcOutput
ENDFUNC
*   ToString()


************************************************************************
*  Clear
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Clear()

THIS.Dispose()
THIS.Init()

ENDFUNC
*   Clear

ENDDEFINE
*EOC wwFileStream 


*************************************************************
DEFINE CLASS wwMemoryStream AS Custom
*************************************************************
*: Author: Rick Strahl
*:         (c) West Wind Technologies, 2012
*:Contact: http://www.west-wind.com
*:Created: 01/05/2012
*************************************************************
#IF .F.
*:Help Documentation
*:Topic:
Class wwMemoryStream

*:Description:

*:Example:

*:Remarks:

*:SeeAlso:


*:ENDHELP
#ENDIF

cOutput = ""
nLength = 0

************************************************************************
*  Destroy
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Destroy()
THIS.Dispose()
ENDFUNC
*   Destroy

************************************************************************
*  Dispose
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Dispose()
this.cOutput = ""
this.nLength = 0
ENDFUNC
*   Dispose

************************************************************************
*  Clear
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Clear()
this.cOutput = ""
this.nLength = 0
ENDFUNC
*   Clear

************************************************************************
*  Write
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Write(lcContent)
this.nLength = this.nLength + LEN(lcContent)
this.cOutput = this.cOutput + lcContent
ENDFUNC
*   Write

************************************************************************
*  WriteLine
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WriteLine(lcContent)
this.Write(lcContent)
this.Write(CRLF)
ENDFUNC
*   WriteLine

************************************************************************
*  WriteFile
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION WriteFile(lcFileName)
this.Write(FILETOSTR( FULLPATH(lcFileName) ))
ENDFUNC
*   WriteFile

************************************************************************
*  ToString()
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ToString()
RETURN this.cOutput
ENDFUNC
*   ToString()

ENDDEFINE
*EOC wwMemoryStream 