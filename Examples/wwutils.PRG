************************************************************************
* FUNCTION Utils
******************
***    Author: Rick Strahl
***            (c) West Wind Technologies, 1995-2011
***   Contact: rstrahl@west-wind.com
***  Modified: 07/10/2011
***  Function: A set of utility classes and functions used by 
***            the various classes and processing code.
*************************************************************************
#INCLUDE FOXPRO.H
#INCLUDE WCONNECT.H

SET PROCEDURE TO wwAPI ADDITIVE
SET PROCEDURE TO wwUtils ADDITIVE

*************************************************************
DEFINE CLASS wwEnv AS Custom
*************************************************************
***  Function: Saves environment settings.
*************************************************************

*** Custom Properties
PROTECTED cSetting,vOldValue

************************************************************************
* wwEnv :: Init
*********************************
***  Function: Saves and restores environment settings
***    Assume: Limited to simple ON/OFF settings
***            Very limited!!! Test carefully.
***      Pass: tcSetting  -   SET value to set
***            tvNewValue -   Value to set to
***    Return:
************************************************************************
FUNCTION Init
LPARAMETERS tcSetting,tvNewValue
THIS.Set(tcSetting, tvNewValue)
ENDFUNC
* Init

************************************************************************
* wwEnv :: Set
*********************************
FUNCTION Set
LPARAMETERS tcSetting,tvNewValue
THIS.cSetting=tcSetting

THIS.vOldValue=SET( tcSetting )

IF TYPE("tvNewValue")="C" AND ;
   INLIST(UPPER(tvNewValue),"ON","OFF") 
   SET &tcSetting &tvNewValue
ELSE
   SET &tcSetting TO (tvNewValue)   
ENDIF

ENDFUNC
* Set

************************************************************************
* wwEnv :: Destroy
*********************************
FUNCTION Destroy
LOCAL lcSetting,lvValue

lcSetting=THIS.cSetting
lvValue=THIS.vOldValue

IF TYPE("lvValue")="C" AND ;
   INLIST(UPPER(lvValue),"ON","OFF") 
   SET &lcSetting &lvValue
ELSE
   SET &lcSetting TO (lvValue)   
ENDIF

ENDFUNC
* Destroy

ENDDEFINE
*EOC wwEnv


#IF .F.
*DEFINE CLASS wwUtils as Custom
#ENDIF



*************************************************************************
****
**** STANDALONE FUNCTIONS
****
*************************************************************************

************************************************************************
FUNCTION OpenExclusive
**********************
***  Modified: 01/27/96
***  Function: Tries to open a table exclusively
***    Assume: Table name can't contain a file name.
***            Returns .F. for other reasons like file !found etc.
***            USES wwEVAL object to test for success
***            Parameters MUST NOT BE LPARAMETERS!!!
***      Pass: lcTable   -  Name of table to open exclusively
***    Return: .T. or .F.
************************************************************************
PARAMETERS lcTable, lcAlias
LOCAL lcOldError, llRetVal, loEval

lcTable=IIF(EMPTY(lcTable),"",lcTable)
lcAlias=IIF(EMPTY(lcAlias),JustStem(lcTable),lcAlias)

IF EMPTY(lcTable)
   RETURN .F.
ENDIF

llRetVal = .t.
TRY 
	*** Use Exclusively to reindex and pack
	IF !USED(lcAlias)
	   USE (lcTable) EXCLUSIVE IN 0 ALIAS (lcAlias)
	ELSE
	   SELE (lcAlias)
	   USE (lcTable) EXCLUSIVE  ALIAS (lcAlias)
	ENDIF
CATCH
   llRetVal = .F.
ENDTRY   

IF !llRetVal
   *** Failed exclusive: try to re-open table as shared
   USE (lcTable) IN 0 ALIAS(lcAlias)
ELSE
   SELE (lcAlias)   
ENDIF   

RETURN llRetVal
*EOP OpenExclusive


************************************************************************
FUNCTION File2Var
******************
***  Function: Takes a file and returns the contents as a string or
***            Takes a string and stores it in a file if a second
***            string parameter is specified.
***      Pass: tcFilename  -  Name of the file
***            tcString    -  If specified the string is stored
***                           in the file specified in tcFileName
***    Return: file contents as a string
************************************************************************
LPARAMETERS tcFileName, tcString, llSharedWrite
LOCAL lcRetVal, lnHandle, lnSize

tcFileName=IIF(EMPTY(tcFileName),"",tcFileName)

IF VARTYPE(tcString) # "C"
   *** Read File to string - if possible use FILETOSTR
   *** since it's native and faster, but we need
   *** to wrap it into an error handler
   lcRetVal=""
   
   TRY
      lcRetVal = FILETOSTR(tcFilename)
   CATCH
   ENDTRY
ELSE
   *** Write a file to string
   tcString=IIF(EMPTY(tcString),"",tcString)
   
   IF !llSharedWrite
      *** Text to File
      lnHandle=FCREATE(tcFileName)
      IF lnHandle=-1
         RETURN .F.
      ENDIF
      =FWRITE(lnHandle,tcString)
      =FCLOSE(lnHandle)
      RETURN .T.
   ELSE
      LOCAL llFailed
      llFailed = .F.
      TRY
         STRTOFILE(tcString,tcFileName)         
      CATCH
         llFailed=.T.
      ENDTRY
      
      IF llFailed
         RETURN .F.
      ENDIF
      RETURN .T.
   ENDIF
ENDIF


IF lcRetVal = "﻿"
	lcRetVal = STRCONV(SUBSTR(lcRetVal,4),11)
ENDIF

RETURN lcRetVal
*EOP File2Var

***********************************************************
FUNCTION LogString(lcOutput, lcFilename) 
*******************
LOCAL lhFile

IF (EMPTY(lcFileName) )
   IF VARTYPE(__LOGPATH) = "C"
   	  lcFileName = __LOGPATH
   ENDIF
   IF EMPTY(lcFileName)
	   lcFilename = SYS(2023) + "\__LogString.txt"
   ENDIF
ENDIF

   
lhFile=FOPEN(lcFileName,2)
IF lhFile = -1 AND !FILE(lcFileName)
   lhFile = FCREATE(lcFilename)
ENDIF   
IF lhFile = -1
   lhFile=FOPEN(lcFileName,2)
   IF lhFile = -1
	   RETURN
   ENDIF
ENDIF   

FSEEK(lhFile,0,2)
FWRITE(lhFile,TRANSFORM(DATETIME()) + " - " + lcOutput + CRLF)
FFLUSH(lhFile)
FCLOSE(lhFile)

ENDFUNC


***********************************************************
FUNCTION CopyObject
*******************
***    Author: Rick Strahl, West Wind Technologies
***            http://www.west-wind.com/
***  Function: Function that copies an object and all of 
***            it child object into a totally new object 
***            reference that is not based on a previous 
***            reference
***    Assume: Support for 1 dimensional arrays only
***      Pass: loObject - Existing object reference to copy
***    Return: New Object reference
***********************************************************
LPARAMETER loInput
LOCAL loObject, laFields[1], lnX, lcField, lcType, llClass, ;
      lnCount, z, lnLength

#DEFINE COPYOBJECT_PROPERTYEXCLUSIONLIST ;
 ",activecontrol,classlibrary,baseclass,comment,controls,objects,controlcount,"+;
 "class,name,parent,parentalias,parentclass,helpcontextid,whatsthishelpid," +;
 "width,height,top,left,tag,picture,onetomany,childalias,childorder,relationalexpr,timestamp_column," 


*** Check if we can instantiate input object class
IF TYPE("loInput.Class") = "C"
   *** If we have a class create it
   loObject = CREATEOBJECT(loInput.CLASS)
ELSE
   *** SCATTER NAME Objects don't have a class
   #IF wwVFPVersion < 8
      loObject = CREATEOBJECT("Relation")
   #ELSE
      loObject = CREATEOBJECT("EMPTY")
   #ENDIF
ENDIF

*** Grab member properties and loop through 'em
lnCount = AMEMBERS(laFields, loInput)
FOR lnX=1 TO lnCount
   lcField = LOWER(laFields[lnX])
   
   *** Ignore stock properties
   IF AT("," + lcField + ",","'" + ;
         COPYOBJECT_PROPERTYEXCLUSIONLIST) > 0
      LOOP
   ENDIF
   
   *** Grab the type
   lcType = TYPE("loInput."+lcField)

   DO CASE
      *** Must check for array properties first
      CASE TYPE([ALEN(loInput.] + lcField + [)]) = "N"
         *** Add property if it doesn't exist
         IF TYPE("loObject." + lcField) = "U"
            #IF wwVFPVersion < 8
                  loObject.ADDPROPERTY(lcField + "[1]")
            #ELSE
                  ADDPROPERTY(loObject,lcField + "[1]")
            #ENDIF 
         ENDIF

         *** Create the array then run through
         *** NOTE: only 1d supported
         lnLength = ALEN(loInput.&lcField)
         DIMENSION loObject.&lcField[lnLength]
         FOR z=1 TO lnLength
            IF TYPE("loInput." +lcField + "[z]")="O"
               loObject.&lcField[z] = ;
                  CopyObject(EVAL( "loInput." + ;
                                   lcField + "[z]"))
            ELSE
               loObject.&lcField[z] = EVAL("loInput." + ;
                                           lcField)
            ENDIF
         ENDFOR
         
      *** Recursive calls for objects
      CASE lcType = "O"
         IF TYPE("loObject." + lcField) = "U"
            #IF wwVFPVersion < 8
               loObject.ADDPROPERTY(lcField)
            #ELSE
               ADDPROPERTY(loObject,lcField)            
            #ENDIF
         ENDIF

         loObject.&lcField = ;
                  CopyObject(EVAL("loInput."+lcField))
      OTHERWISE
         *** Check if property exists
         IF TYPE("loObject." + lcField) = "U"
            #IF wwVFPVersion < 8
               loObject.ADDPROPERTY(lcField)
            #ELSE
               ADDPROPERTY(loObject,lcField)            
            #ENDIF
         ENDIF

         *** Straight Assignment
         loObject.&lcField = EVAL("loInput." + lcField)
   ENDCASE
ENDFOR

RETURN loObject
* EOF CopyObject

#IF wwVFPVersion > 7
************************************************************************
* wwUtils :: CopyObjectProperties
****************************************
***  Function: Copies properties from one object to another
***    Assume: Only copies single level
***      Pass:
***    Return:
************************************************************************
FUNCTION CopyObjectProperties(loInput, loObject, lnObjectStructureObject)

IF ( EMPTY(lnObjectStructureObject) )
   lnObjectStructureObject = 1
ENDIF

#DEFINE COPYOBJECT_PROPERTYEXCLUSIONLIST ;
 ",activecontrol,classlibrary,baseclass,comment,controls,objects,controlcount,"+;
 "class,name,parent,parentalias,parentclass,helpcontextid,whatsthishelpid," +;
 "width,height,top,left,tag,picture,onetomany,childalias,childorder,relationalexpr,timestamp_column," 

*** Grab member properties and loop through 'em
IF lnObjectStructureObject = 1
   lnCount = AMEMBERS(laFields, loInput)
ELSE
   lnCount = AMEMBERS(laFields, loObject)
ENDIF

FOR lnX=1 TO lnCount
   lcField = LOWER(laFields[lnX])
   
   *** Ignore stock properties
   IF AT("," + lcField + ",","'" + ;
         COPYOBJECT_PROPERTYEXCLUSIONLIST) > 0
      LOOP
   ENDIF
   
   *** Grab the type
   lcType = TYPE("loInput."+lcField)

   IF lcType = "O" OR lcType = "U" OR ;
      TYPE([ALEN(loInput.] + lcField + [)]) = "N"
      * TYPE("loInput."+lcField,1) = "A"
      loop
   ENDIF
   
   *** Check if property exists
   IF lnObjectStructureObject = 1
      IF TYPE("loObject." + lcField) = "U"
         LOOP
      ENDIF
   ELSE
      IF TYPE("loInput." + lcField) = "U"
         LOOP
      ENDIF
   ENDIF

   *** If so assign it
   TRY
      loObject.&lcField = EVAL("loInput." + lcField)
   CATCH
      *** Just ignore error
   ENDTRY
   
   DO CASE
      *** Must check for array properties first
      CASE TYPE([ALEN(loInput.] + lcField + [)]) = "N"

         ** Do nothing
                 
      *** Recursive calls for objects
      CASE lcType = "O"
      *** Do nothing for now
*!*            IF TYPE("loObject." + lcField) = "U"
*!*               #IF wwVFPVersion < 8
*!*                  loObject.ADDPROPERTY(lcField)
*!*               #ELSE
*!*                  ADDPROPERTY(loObject,lcField)            
*!*               #ENDIF
*!*            ENDIF

*!*               loObject.&lcField = ;
*!*                      CopyObject(EVAL("loInput."+lcField))
      OTHERWISE
   ENDCASE
ENDFOR

ENDFUNC
#ENDIF

************************************************************************
* GetRegExObject
****************************
***  Function: Creates an instance of a RegEx object and caches it
****           in a public variable so it can be reused.
***    Return: Object or .NULL. if it can't be created
************************************************************************
FUNCTION GetRegExObject

IF VARTYPE(__REGEX) != "O"
   PUBLIC __REGEX 
   __REGEX = CREATEOBJECT("VBScript.RegExp")
   __REGEX.Global = .T.
*   __REGEX.IgnoreCase = .T.
ENDIF

RETURN __REGEX
*  wwUtils ::  GetRegexObject

************************************************************************
* wwUtils ::  GetwwRegexObject
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetwwRegexObject()

IF VARTYPE(__WWREGEX) != "O"
   PUBLIC __WWREGEX 
   LOCAL lcRegEx
   lcRegEx = "wwRegEx.prg"
   __WWREGEX = NEWOBJECT("wwRegEx",lcRegEx)
   __WWREGEX.Global = .T.
*   __REGEX.IgnoreCase = .T.
ENDIF

RETURN __WWREGEX
*  wwUtils ::  GetwwRegexObject

************************************************************************
* FUNCTION GetXmlDomObject
****************************
***  Function: Creates an instance of a Xml DOM object and caches it
****           in a public variable so it can be reused.
***    Return: Object or .NULL. if it can't be created
************************************************************************
FUNCTION GetXmlDomObject(llAsync)

IF VARTYPE(__XMLDOM) != "O"
   PUBLIC __XMLDOM 
   __XMLDOM = CREATEOBJECT(XML_XMLDOM_PROGID)
   __XMLDOM.Async = llAsync
ENDIF

RETURN __XMLDOM


************************************************************************
FUNCTION CursorToObjectArray
****************************
***  Function: Creates an array of objects of the currently open
***            cursor. The result is an object that has two properties
***            the count and the array of objects that contain the
***            records. The object has properties for each field
***            value (except general fields)
*** Parameter: lcClass  -  Optional class that has aRows and nCount 
***                        members         
***    Return: Object or .NULL. if no records in cursor
************************************************************************
LPARAMETER lcObjName
LOCAL lnX, laLItems[1], lnSize

*** Create object and set count and array properties
IF !EMPTY(lcObjName)
      loResult = CREATE(lcObjName)
ELSE   
   #IF WWVFPVERSION < 8
   	loResult = CREATE("RELATION")
   	loResult.ADDPROPERTY("nCount")
   	loResult.AddProperty("aRows(1)",1)
   #ELSE
      loResult = CREATE("EMPTY")
      ADDPROPERTY(loResult,"nCount",0)
      ADDPROPERTY(loResult,"aRows(1)",1)
   #ENDIF
ENDIF

lnX = 0
COUNT TO lnSize

*** Handle empty cursor
IF lnSize = 0
   RETURN NULL   
ENDIF   

DIMENSION loResult.aRows[lnSize]
SCAN
   lnX=lnX+1
   SCATTER NAME loResult.aRows[lnX] MEMO
ENDSCAN

loResult.nCount = lnX

RETURN loResult
*EOP CursorToObjectArray

************************************************************************
FUNCTION CacheFile
******************
***  Function: Caches read from disk in a cursor. Use like File2Var.
***    Assume: Using llCheckDate can worsen performance considerably
***      Pass: lcFileName       -   Name of the file to read/cache
***            lnRefreshSeconds -   If the file cached is older than
***                                 the number of seconds specified
***                                 here it's reloaded. Use 0 to 
***                                 never reload.
***            llMergeText       -  if .T. merges ASP style tags
***                                 on the content from the file
***                                 and stores result in the cache
***    Return: String of the file or "" 
************************************************************************
LPARAMETERS lcFileName, lnRefreshSeconds, llMergeText
LOCAL lcOutput, lnHandle, __loEval

lcAlias = ALIAS()

IF USED("wwFileCache")
   SELECT wwFileCache
ELSE 
   CREATE CURSOR wwFileCache ;
	  (FileName C(120),;
	   TimeRead T,;
	   Content M)
   INDEX ON FileName Tag FileName
ENDIF   

lcFileName = LOWER(lcFileName)

IF SEEK(lcFileName)
    *** Reread the file if timed out
	IF lnRefreshSeconds > 0 AND;
	   wwFileCache.TimeRead < DATETIME() - lnRefreshSeconds

		 *** File's timed out - re-read and update
	     lcOutput = File2Var(lcFileName)
	     
	     IF llMergeText
	         lcOutput = MergeText(@lcOutput)
	     ENDIF

	     REPLACE Content with lcOutput,;
		          TimeRead with DATETIME()
	ENDIF

    lcOutput = wwFileCache.Content
ELSE
     *** File's not in the cache yet - insert it
     lcOutput = File2Var(lcFileName)
     
     IF llMergeText
         lcOutput = MergeText(@lcOutput)
     ENDIF
     
     INSERT INTO wwFileCache (FileName, TimeRead, Content) ;
	          VALUES (lcFileName,DATETIME(),lcOutput)
ENDIF          

IF !EMPTY(lcAlias)
  SELECT (lcAlias)
ENDIF

RETURN lcOutput
ENDFUNC


************************************************************************
FUNCTION WrCursor
******************
***  Function: Creates a Writable cursor from a SELECTed cursor.
***            The original cursor is closed after completion.
***
***    Assume: No checks for valid cursor are performed! Make sure
***            you have a cursor and not a buffered table image!
***            CURSOR MUST BE CURRENTLY SELECTED!
***
***      Pass: pcNewName  -  New Alias name for the cursor
***
***    Return: nothing
************************************************************************
PARAMETER pcNewName
PRIVATE lcOldAlias

*** work with the current Cursor
lcOldAlias=ALIAS()

*** Select the cursor and get file name
SELE (lcOldAlias)
lcDBF=DBF(lcOldAlias)

*** Make sure things are closed in the new alias
IF USED(pcNewName)
   USE IN (pcNewName)
ENDIF

USE (lcDBF) AGAIN ALIAS (pcNewName) IN 0

USE IN (lcOldAlias)

SELE (pcNewName)

RETURN 


************************************************************************
*  IsNullOrEmpty
****************************************
***  Function: Returns whether a value is null or Empty
***    Assume:
***      Pass: lvValue - value 
***    Return: .T. or .F.
************************************************************************
FUNCTION IsNullOrEmpty(lvValue,llCheckArray)

IF llCheckArray
	lcType = TYPE("lvValue")
	IF lcType = "A"
	   RETURN .T.
	ENDIF
ELSE	
	lcType = VARTYPE(lvValue)
ENDIF	

IF lcType = "X" OR lcType = "U"
   RETURN .T.
ENDIF   

IF INLIST(lcType,"C","M","N","Y","D","T","L") AND EMPTY(lvValue)
   RETURN .T.
ENDIF

RETURN .F.  
ENDFUNC
*   IsNullOrEmpty


************************************************************************
*  StringFormat
****************************************
***  Function: Allows for {0} value replacement in string 
***    Assume: 
***      Pass:
***    Return:
************************************************************************
FUNCTION StringFormat(lcFormat,parm1,parm2,parm3,parm4,parm5,parm6,parm7,parm8,parm9,parm10)
LOCAL lnParmCount, lnX

lnParmCount = PARAMETERS() - 1

FOR lnX = 1 TO lnParmCount
	lcFormat = STRTRAN(lcFormat,;
					"{" + TRANSFORM(lnX -1) + "}",;
			        TRANSFORM( EVALUATE("Parm" + TRANSFORM(lnX)) ))
ENDFOR

RETURN lcFormat
ENDFUNC
*   StringFormat

************************************************************************
FUNCTION Extract
******************
***  Function: Extracts a text value between two delimiters
***    Assume: Delimiters case insensitive
***            The first instance only is retrieved. Idea is
***            to translate the delims as you go...
***      Pass: lcString   -  Entire string
***            lcDelim1   -  The starting delimiter
***            lcDelim2     -  Ending delimiter
***            lcDelim3     -  Alternate ending delimiter
***            llEndOk     -  End of line is OK
***            llIncludeDelims - returns match with delimiters
***    Return: Text between delimiters or ""
*************************************************************************
PARAMETERS lcString,lcDelim1,lcDelim2,lcDelim3, llEndOk, llIncludeDelims
PRIVATE lnX,lnLocation,lcRetVal,lcChar,lnNewString,lnEnd

#IF wwVFPVersion > 8
   IF EMPTY(lcDelim3) 
      RETURN STREXTRACT(lcString,lcDelim1,lcDelim2,1,1 + IIF(llendOk,2,0) + IIF(llIncludeDelims,4,0))
   ENDIF
#ENDIF

lcDelim1=IIF(LEN(lcDelim1)=0,",",lcDelim1)
lcDelim2=IIF(LEN(lcDelim2)=0,"z!x",lcDelim2)
lcDelim3=IIF(EMPTY(lcDelim3),"z!x",lcDelim3)

lnLocation=ATC(lcDelim1,lcString)
IF lnLocation=0
   RETURN ""
ENDIF

IF llIncludeDelims
   *** Crate a new string of remaining text
   lcNewString=lcString

   lnEnd=ATC(lcDelim2,lcNewString)
   IF lnEnd>0
      RETURN SUBSTR(lcNewString,lnLocation,(lnEnd + LEN(lcDelim2)) - lnLocation)
   ENDIF   
   *!*   IF lnEnd = 0
   *!*      *** Empty Delimited string
   *!*      RETURN ""
   *!*   ENDIF
      
   lnEnd=ATC(lcDelim3,lcNewString)
   IF lnEnd>0
      RETURN SUBSTR(lcNewString,1,(lnEnd + LEN(lcDelim2)) - lnLocation)
   ENDIF   

   IF llEndOk
     *** Return to the end of the line
     RETURN lcNewString
   ENDIF
ELSE
   lnLocation=lnlocation+len(lcDelim1)  && -1 

   *** Crate a new string of remaining text
   lcNewString=SUBSTR(lcString,lnLocation)

   lnEnd=ATC(lcDelim2,lcNewString)
   IF lnEnd>0
      RETURN SUBSTR(lcNewString,1,lnEnd-1)
   ENDIF   
   *!*   IF lnEnd = 0
   *!*      *** Empty Delimited string
   *!*      RETURN ""
   *!*   ENDIF
      
   lnEnd=ATC(lcDelim3,lcNewString)
   IF lnEnd>0
      RETURN SUBSTR(lcNewString,1,lnEnd-1)
   ENDIF   

   IF llEndOk
     *** Return to the end of the line
     RETURN SUBSTR(lcNewString,1)
   ENDIF
ENDIF

RETURN ""
*EOP RetValue


************************************************************************
FUNCTION GetAttribute
*************************
***  Function: Retrieves an HTML/XML attribute from an element string
***      Pass: lcElementXml - full Element HTML/XML string
***            lcAttribute  - the name of hte attribute
***    Return: the content of the tag or ""
*************************************************************************
LPARAMETERS lcElementXml,lcAttribute
LOCAL lnAt, lnAt2, lcSub

lnAt = ATC(lcAttribute+"=",lcElementXml)
IF lnAt = 0
   RETURN ""
ENDIF

lcSub = SUBSTR(lcElementXml,lnAt + LEN(lcAttribute)+1)
IF lcSub = ["] 
   lnAt = AT(["],lcSub,2)
   IF lnAt = 0
      RETURN ""  
   ENDIF
   RETURN SUBSTR(lcSub,2,lnAt-2)
ENDIF 
IF lcSub = ['] 
   lnAt = AT(['],lcSub,2)
   IF lnAt = 0
      RETURN ""  
   ENDIF
   RETURN SUBSTR(lcSub,2,lnAt-2)
ENDIF 
  
*** No quotes around - must look for space or >
lnAt = AT([ ],lcSub)
lnAt2 = AT([>],lcSub)

IF lnAt > 0 AND lnAt < lnAt2
   RETURN SUBSTR(lcSub,1,lnAt-1)
ENDIF   

RETURN SUBSTR(lcSub,1,lnAt2-1)




************************************************************************
* wwUtils :: ReplaceText
****************************************
***  Function: Replaces text between a set of delimiters with 
***            a new string leaving the delimiters intact
***    Assume: The delimited block MUST have at least 1 character
***            in it even if it is blank or a CRLF
***      Pass: lcSource - Full Source String
***            lcStart  - Start delimiter
***            lcEnd    - End Delimiter
***            lcReplace -  Text to replace between delimiters
***    Return: Updated String  
************************************************************************
FUNCTION ReplaceText
LPARAMETERS lcSource, lcStart, lcEnd, lcReplace

lcExtract = Extract(lcSource,lcStart,lcEnd,.f.,.T.)
  
RETURN STRTRAN(lcSource,lcStart + lcExtract + lcEnd,lcStart + lcReplace + lcEnd)
ENDFUNC

************************************************************************
* wwUtils ::  ReplaceTextAndDelimiters
****************************************
***  Function:
***    Assume:
***      Pass: 
***    Return:
************************************************************************
FUNCTION ReplaceTextAndDelimiters
LPARAMETERS lcSource, lcStart, lcEnd, lcReplace, llAll

IF llAll
    LOCAL lcNewText
	lcNewText = ""
	DO WHILE .T.
		lcNewText = ReplaceTextAndDelimiters(lcSource,lcStart,lcEnd,lcReplace)
		IF lcSource = lcNewText
		   lcSource = lcNewText
		   EXIT
		ENDIF
		lcSource = lcNewText
	ENDDO	
   RETURN lcSource
ENDIF

lcExtract = Extract(lcSource,lcStart,lcEnd,,.T.)
  
RETURN STRTRAN(lcSource,lcStart + lcExtract + lcEnd,lcReplace)
ENDFUNC
*  wwUtils ::  ReplaceTextAndDelimiters


************************************************************************
* wwUtils :: TrimWhiteSpace
****************************************
***  Function: Trims Whitespace from the end of a string
***    Assume:
***      Pass: lcChars  -  Optional: Characters to consider white space
***    Return: Trimmed String
************************************************************************
FUNCTION TrimWhiteSpace(lcString, lcChars)
LOCAL lnX

#IF WWVFPVERSION > 8
IF EMPTY(lcChars)
   RETURN TRIM(lcString,0," ",CHR(13),CHR(10),CHR(9))
ENDIF
#ENDIF

IF EMPTY(lcChars)
  lcChars = " " + CHR(13) + CHR(10) + CHR(9)
ENDIF

FOR lnX = LEN(lcString) TO 1  STEP -1
   IF AT(SUBSTR(lcString,lnX,1),lcChars) < 1
      RETURN SUBSTR(lcString,1,lnX)
   ENDIF
ENDFOR

RETURN ""
ENDFUNC
*  wwUtils :: TrimWhiteSpace

#IF wwVFPVERSION > 8
************************************************************************
* wwUtils :: AllTrimWhiteSpace
****************************************
***  Function: Trims Whitespace from the beginning and end of a string
***    Assume:
***      Pass: lcChars  -  Optional: Characters to consider white space
***    Return: Trimmed String
************************************************************************
FUNCTION AllTrimWhiteSpace(lcString)
RETURN ALLTRIM(lcString,0," ",CHR(13),CHR(10),CHR(9))
ENDFUNC
*  wwUtils :: AllTrimWhiteSpace
#ENDIF


************************************************************************
*  StartsWith
****************************************
***  Function: Compares a source string and a compare string
***            to see if source string starts with the compare string
***    Assume:
***      Pass: lcSourceString   -  Base string to check
***            lcCompare        -  Start string to check against
***    Return: .T. or .F. 
************************************************************************
FUNCTION StartsWith(lcSourceString,lcCompare, llCaseInSensitive)

IF EMPTY(lcCompare)
	IF EMPTY(lcSourceString)
	   RETURN .T.
	ENDIF
	RETURN .F.
ENDIF

IF !llCaseInSensitive
	IF  LEFT(lcSourceString,LEN(lcCompare)) == lcCompare
		RETURN .T.
	ENDIF
ELSE
	IF  LEFT(LOWER(lcSourceString),LEN(lcCompare)) == LOWER(lcCompare)
		RETURN .T.
	ENDIF
ENDIF

RETURN .F.	
ENDFUNC
*   StartsWith

************************************************************************
*  EndsWith
****************************************
***  Function: Compares a source string and a compare string
***            to see if source string ends with the compare string
***    Assume:
***      Pass: lcSourceString   -  Base string to check
***            lcCompare        -  Start string to check against
***    Return: .T. or .F. 
************************************************************************
FUNCTION EndsWith(lcSourceString,lcCompare,llCaseInSensitive)

IF EMPTY(lcCompare)
	IF EMPTY(lcSourceString)
	   RETURN .T.
	ENDIF
	RETURN .F.
ENDIF

IF !llCaseINSensitive
	IF  RIGHT(lcSourceString,LEN(lcCompare)) == lcCompare
		RETURN .T.
	ENDIF
ELSE
	IF  RIGHT(LOWER(lcSourceString),LEN(lcCompare)) == LOWER(lcCompare)
		RETURN .T.
	ENDIF
ENDIF

RETURN .F.	
ENDFUNC
*   EndsWith

************************************************************************
* wwUtils :: UnicodeToUtf8
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION UnicodeToUtf8(lcUnicodeInputString)

DECLARE INTEGER WideCharToMultiByte IN kernel32; 
    INTEGER   CodePage,; 
    INTEGER   dwFlags,; 
    STRING    lpWideCharStr,; 
    INTEGER   cchWideChar,; 
    STRING  @ lpMultiByteStr,; 
    INTEGER   cbMultiByte,; 
    STRING    lpDefaultChar,; 
    INTEGER   lpUsedDefaultChar 
      
   
*CP_UTF8 = 65001
* CP_ACP =0

LOCAL lnNewSize, lcBuffer,lnAnsiSize, lnOldSize
lnNewSize =LEN(lcUnicodeInputString) * 3
lcBuffer = REPLICATE(CHR(0),lnNewSize)
lnOldSize = LEN(lcUniCodeInputString)/2

lnAnsiSize = WideCharToMultiByte(65001, 0, lcUnicodeInputString,lnOldSize, ;
                                 @lcBuffer,lnNewSize, 0, 0)
                                 
IF lnAnsiSize =0 
   RETURN ""
ENDIF

RETURN SUBSTR(lcBuffer,1,lnAnsiSize)
ENDFUNC
*  wwUtils :: UnicodeToUtf8

************************************************************************
FUNCTION Path
******************
***  Function: Adds or deletes items from the path string
***      Pass: pcPathName   -   Filename
***            pcMethod     -   *"ADD","DELETE"
***    Return: New Path or ""
************************************************************************
PARAMETERS pcPath,pcMethod
LOCAL lcOldPath

IF VARTYPE(pcMethod) # "C"
   pcMethod = "ADD"
ENDIF

IF EMPTY(pcPath)
   RETURN
ENDIF

pcPath=ADDBS(LOWER(TRIM(pcPath)))

#IF wwVFPVersion > 8
IF pcMethod = "ADD"
   IF !EMPTY(pcPath)	
	   SET PATH TO (pcPath) Additive
   	   RETURN SET("PATH")
   ENDIF
ENDIF
#ENDIF   

lcOldPath=LOWER(SET("PATH"))

IF pcMethod="ADD"
	    IF EMPTY(pcPath) .OR. ;
	       !Directory(pcPath,1)
	       RETURN ""
	   ENDIF
	   IF AT(";" + pcPath + ";" ,";" + lcOldPath + ";")>0
	      RETURN ""
	   ENDIF
	   lcOldPath=lcOldPath+";"+pcPath
ELSE
   IF AT(";" + pcPath + ";" ,";" + lcOldPath + ";") < 1
      RETURN ""
   ENDIF
   lcOldPath=STRTRAN(lcOldPath + ";" ,";" +pcPath+";",";")
   lcOldPath = SUBSTR(lcOldPath,1,LEN(lcOldPath)-1)
ENDIF   

SET PATH TO &lcOldPath

RETURN lcOldPath
*EOP PATH


************************************************************************
FUNCTION DomainName
*******************
***  Modified: 04/13/96
***  Function: Retrieves a Domain name from an URL
***    Assume: URL starts with http:// - // required!
***      Pass: lcUrl         -  URL to retrieve name from
***            llNoStripWWW  -  Don't strip www.
***    Return: Domain Name or ""
*************************************************************************
LPARAMETER lcUrl, llNoStripWWW
lcText=STRTRAN(EXTRACT(lower(lcUrl),"//","/"," "),"/","")
IF !llNoStripWWW
  lcText=STRTRAN(lcText,"www.","")
ENDIF
RETURN PADR(lcText,50)


************************************************************************
FUNCTION PropertyDump
**********************
***  Function: Dumps all of an objects properties to a string separated
***            by Carriage Returns. Note Long strings will be truncated
***            at 80 characters
***      Pass: loObject  -  Object to work with
***    Return: string of key value pairs
*************************************************************************
LPARAMETER loObject, lnMode
LOCAL lnX, lnCount, lcOutput

IF EMPTY(lnMode)
	lnMode = 1
ENDIF

lnCount = AMEMBERS(laFields, loObject,lnMode)
lcOutput = ""
FOR lnX=1 to lnCount
   lcType = TYPE("loObject."+laFields[lnX,1])
   IF ATC(lcType,"UO") = 0
      lvValue = EVAL("loObject."+laFields[lnX,1])
      IF lcType="C" AND LEN(lvValue) > 80
         lvValue = LEFT(lvValue,80)
      ENDIF
      lcOutput = lcOutput  + CHR(13) +  CHR(10) + lower(laFields[lnX,1]) + " = " + ALLTRIM(TRANSFORM(lvValue,""))
   ELSE
      lcOutput = lcOutput + CHR(13) + CHR(10) +Lower(laFields[lnX,1]) + " = " + IIF(TYPE("loObject."+laFields[lnX,1])="O","Object","NULL")
   ENDIF
ENDFOR

RETURN lcOutput
* EOF PropertyDump


************************************************************************
FUNCTION FixPreTags
********************
***  Function: Fixes <PRE> tags in HTML pages so that they
***            display properly and can be cut and pasted
***            as code. Replaces <p> and <br> tags with
***            real carriage returns.
***      Pass: lcHTML  -   HTML to fix
***    Return: Fixed HTML
*************************************************************************
LPARAMETER lcHTML, lnColWidth
LOCAL lcPre, lcFixed, lcPrecount, lnAt1, lnAt2

*** Fix up <Pre> formatted text. Use plain Returns
*** so the code can be pasted properly
lnPrecount = 1
DO WHILE .T.
   lnAt1 = ATC("<pre",lcHTML,lnPreCount)
   lnAt2 = ATC("/pre>",lcHTML,lnPreCount)
   IF lnAt1 = 0 or lnAt2 = 0
      EXIT
   ENDIF
   lcPre = substr(lcHTML,lnAt1,lnAt2 - lnAt1)
   
   lcFixed = STRTRAN(lcPre,"<p>",CHR(13)+CHR(10)+CHR(13)+CHR(10))
   lcFixed = STRTRAN(lcFixed,"<br>" + CHR(13),CHR(13)+CHR(10))
   lcFixed = STRTRAN(lcFixed,"<br>",CHR(13)+CHR(10))

   lcHTML = STRTRAN(lcHTML,lcPre,lcFixed)
   lnPreCount = lnPreCount + 1
ENDDO

RETURN lcHTML


************************************************************************
Function FixHTMLForDisplay
************************************************************************
#IF .F.
*:Help Documentation
*:Topic:
wwUtils::FixHTMLForDisplay

*:Description:
This method fixes up HTML for display. Takes HTML and XML tags and converts
them to HTML displayable characters (&lt; for < for example).

*:Parameters:
<<b>>lcHTML<</b>>
HTML to fix up.

*:Returns:
Fixed up HTML
*:ENDHELP
#ENDIF
************************************************************************
LPARAMETER lcHTML

lcHTML = STRTRAN(lcHTML,"<","&lt;")
lcHTML = STRTRAN(lcHTML,">","&gt;")
lcHTML = STRTRAN(lcHTML,["],"&quot;")
lcHTML = STRTRAN(lcHTML,[& ],"&amp; ")

RETURN lcHTML

************************************************************************
* EncodeHtml
****************************************
***  Function: Encodes common HTML characters in string so they 
***            display safely.
***    Assume: Same as FixHtmlForDisplay - easier name
***      Pass: lcHtml 
***    Return:
************************************************************************
FUNCTION EncodeHtml(lcHtml)

lcHTML = STRTRAN(lcHTML,"<","&lt;")
lcHTML = STRTRAN(lcHTML,">","&gt;")
lcHTML = STRTRAN(lcHTML,["],"&quot;")
lcHTML = STRTRAN(lcHTML,[& ],"&amp; ")

RETURN lcHTML
ENDFUNC
*  EncodeHtml

************************************************************************
*  DecodeHtml
****************************************
***  Function: Decodes an Html Encoded string back into plain text
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION DecodeHtml(lcHtml)

lcHTML = STRTRAN(lcHTML,"&lt;","<")
lcHTML = STRTRAN(lcHTML,"&gt;",">")
lcHTML = STRTRAN(lcHTML,"&quot;",["])
lcHTML = STRTRAN(lcHTML,"&amp;",[&])

RETURN lcHtml
ENDFUNC
* DecodeHtml


************************************************************************
FUNCTION DisplayMemo
**************************
***  Function: Fixes linebreaks into HTML breaks <br> and <p>
***      Pass: lcHTML  -   HTML to fix
***    Return: Fixed HTML
*************************************************************************
LPARAMETER lcHTML

lcHTML = STRTRAN(lcHTML,CHR(13)+CHR(10),CHR(13))
lcHTML = STRTRAN(lcHTML,CHR(10),CHR(13))
lcHTML = STRTRAN(lcHTML,CHR(13)+CHR(13),"<p>")

RETURN STRTRAN(lcHTML,CHR(13),"<br>")
* FixHTMLLineBreaks

************************************************************************
FUNCTION TextAbstract(lcText, lnLength)
****************************************
***  Function: Creates an abstract string of a certain length
***            and adds ... if truncated
***    Assume: text truncates to the nearest space prior to length
***      Pass:
***    Return:
************************************************************************

IF LEN(lcText) <= lnLength
   RETURN lcText
ENDIF
      
lcText = LEFT(lcText,lnLength)

RETURN LEFT(lcText,RAT(" ",lcText)-1) + "..."
ENDFUNC
*  wwUtils :: TextAbstract


*!*	************************************************************************
*!*	FUNCTION ErrorDisplay
*!*	**********************
*!*	***  Function:
*!*	***    Assume:
*!*	***      Pass:
*!*	***    Return:
*!*	*************************************************************************
*!*	LPARAMETER  lnError, lcMethod, lnLine, lcClass
*!*	LOCAL lcText, lcTitle

*!*	lnError=IIF(EMPTY(lnError),0,lnError)
*!*	lcMethod=IIF(EMPTY(lcMethod),"Unknown",lcMethod)
*!*	lcClass=IIF(EMPTY(lcClass),"Unknown",lcClass)
*!*	lnLine=IIF(EMPTY(lnLine),0,lnLine)

*!*	lcTitle = lcClass + "::" + lcMethod
*!*	lcText = "An error occurred in a class method:" +CHR(13)+CHR(13)+;
*!*	         "Error No:		" + LTRIM(STR(lnError)) + CHR(13) + ;
*!*	         "Error Msg:	"+Message()+CHR(13) + ;
*!*	         "Code:		" + Message(1) + CHR(13) + ;
*!*	         "Line: 		" + LTRIM(STR(lnLine)) + CHR(13)+CHR(13)+;
*!*	         "Do you want to step into the Error Code?"

*!*	lnResult = MessageBox(lcText, 48+4,lcTitle)         
*!*	IF lnResult = 6
*!*	   RETURN .T.
*!*	ENDIF
*!*	   
*!*	RETURN .F. 
*!*	* EOF ErrorDisplay


****************************************************
FUNCTION GoUrl
******************
***    Author: Rick Strahl
***            (c) West Wind Technologies, 1996
***   Contact: rstrahl@west-wind.com
***  Modified: 03/14/96
***  Function: Starts associated Web Browser
***            and goes to the specified URL.
***            If Browser is already open it
***            reloads the page.
***    Assume: Works only on Win95 and NT 4.0
***      Pass: tcUrl  - The URL of the site or
***                     HTML page to bring up
***                     in the Browser
***    Return: 2  - Bad Association (invalid URL)
***            31 - No application association
***            29 - Failure to load application
***            30 - Application is busy 
***
***            Values over 32 indicate success
***            and return an instance handle for
***            the application started (the browser) 
****************************************************
LPARAMETERS tcUrl, tcAction, tcDirectory, tcParms

IF EMPTY(tcUrl)
   RETURN -1
ENDIF
IF EMPTY(tcAction)
   tcAction = "OPEN"
ENDIF
IF EMPTY(tcDirectory)
   tcDirectory = SYS(2023) 
ENDIF

DECLARE INTEGER ShellExecute ;
    IN SHELL32.dll ;
    INTEGER nWinHandle,;
    STRING cOperation,;
    STRING cFileName,;
    STRING cParameters,;
    STRING cDirectory,;
    INTEGER nShowWindow
IF EMPTY(tcParms)
   tcParms = ""
ENDIF

DECLARE INTEGER FindWindow ;
   IN WIN32API ;
   STRING cNull,STRING cWinName

RETURN ShellExecute(FindWindow(0,_SCREEN.caption),;
                    tcAction,tcUrl,;
                    tcParms,tcDirectory,1)


************************************************************************
FUNCTION ShowHTML
*****************
***  Function: Takes an HTML string and displays it in the default
***            browser. 
***    Assume: Uses a file to store HTML temporarily.
***            For this reason there may be concurrency issues
***            unless you change the file for each use
***      Pass: lcHTML       -   HTML to display
***            lcFile       -   Temporary File to use (Optional)
***            loWebBrowser -   Web Browser control ref (Optional)
************************************************************************
LPARAMETERS lcHTML, lcFile, loWebBrowser

lcHTML=IIF(EMPTY(lcHTML),"",lcHTML)
lcFile=IIF(EMPTY(lcFile),SYS(2023)+"\ww_HTMLView.htm",lcFile)

File2Var(lcFile,lcHTML)

IF TYPE("loWebBrowser") = "O"
   loWebBrowser.Navigate(lcFile)
ELSE
   IF TYPE("_oscreenx") = "O"
      _oscreenx.Navigate(lcFile)
   ENDIF
*!*	   IF lower(JUSTEXT(lcFile)) = "txt"
*!*	      MODI COMM (lcFile) IN MACDESKTOP
*!*	   ELSE
      =GoUrl(lcFile)
*!*	   ENDIF
ENDIF   

RETURN
*EOP ShowHTML

************************************************************************
FUNCTION ShowXML
*****************
***  Function: Takes an XML string and displays it in the default
***            browser. 
***    Assume: Uses a file to store HTML temporarily.
***            For this reason there may be concurrency issues
***            unless you change the file for each use
***      Pass: lcHTML       -   HTML to display
***            lcFile       -   Temporary File to use (Optional)
***            loWebBrowser -   Web Browser control ref (Optional)
************************************************************************
LPARAMETERS lcHTML, lcFile, loWebBrowser
IF EMPTY(lcFile)
  lcFile=IIF(EMPTY(lcFile),SYS(2023)+"\ww_HTMLView.xml",lcFile)
ENDIF
ERASE (lcFile) 
RETURN ShowHTML(lcHTML,lcFile,loWebBrowser)

************************************************************************
FUNCTION ShowText
*****************
***  Function: Takes an XML string and displays it in the default
***            browser. 
***    Assume: Uses a file to store HTML temporarily.
***            For this reason there may be concurrency issues
***            unless you change the file for each use
***      Pass: lcHTML       -   HTML to display
***            lcFile       -   Temporary File to use (Optional)
***            loWebBrowser -   Web Browser control ref (Optional)
************************************************************************
LPARAMETERS lcHTML, lcFile, loWebBrowser
IF EMPTY(lcFile)
  lcFile=IIF(EMPTY(lcFile),SYS(2023)+"\ww_HTMLView.txt",lcFile)
ENDIF

IF VARTYPE(loWebBrowser) = "C" and loWebBrowser = "MODI"
   FILE2VAR(lcFile,lcHTML)
   MODIFY COMMAND (lcFile)
   RETURN
ENDIF   

RETURN ShowHTML(lcHTML,lcFile,loWebBrowser)


************************************************************************
FUNCTION StripHTML
*******************
***  Function: Removes HTML tags from the passed text and converts
***            it to plain text. Note formatting is totally removed!
***    Assume: only <br> and <p> are translated
***            any < or > in the HTML besides tags will break this
***            function.
***      Pass: lcText  -   HTML Text to strip
***            lcLTag  -   Left Tag value ("<")
***            lcRTag  -   Right Tag Value (">")
***    Return: Stripped HTML text
*************************************************************************
LPARAMETER lcHTMLText, lcLTag, lcRTag

lcLTag=IIF(EMPTY(lcLTag),"<",lcLTag)
lcRTag=IIF(EMPTY(lcRTag),">",lcRTag)

IF ATC(lcLTag,lcHTMLText) = 0
   RETURN lcHTMLText
ENDIF

*** Start by breaking line breaks
lcHTMLText = STRTRAN(lcHTMLText,lcLTag + "BR" + lcRTag,CRLF)
lcHTMLText = STRTRAN(lcHTMLText,lcLTag + "P" + lcRTag,CRLF+CRLF)
lcHTMLText = STRTRAN(lcHTMLText,lcLTag + "br" + lcRTag,CRLF)
lcHTMLText = STRTRAN(lcHTMLText,lcLTag + "p" + lcRTag,CRLF+CRLF)
lcHTMLText = STRTRAN(lcHTMLText,"&nbsp;"," ")

lcExtract = "x"   
DO WHILE !EMPTY(lcExtract)
   lcExtract = STREXTRACT(lcHtmlText,lcLTag,lcRTag,1)
   **Extract(lcHTMLText,lcLTag,lcRTag)
			   
   IF EMPTY(lcExtract)
      EXIT
   ENDIF
   lcHTMLText = STRTRAN(lcHTMLText,lcLTag+lcExtract+lcRTag,"")
ENDDO

lcHTMLText = STRTRAN(lcHTMLText,"&lt;","<")
lcHTMLText = STRTRAN(lcHTMLText,"&gt;",">")

RETURN lcHTMLText

************************************************************************
FUNCTION HTMLColor
*********************************
***  Function: Converts a FoxPro Color to an HTML Hex color value
***      Pass: lnRGBColor   -  FoxPro RGB color number - RGB(255,255,255)
***            llNoOutput
***    Return: Hex HTML Color String "#FFFFFF"
************************************************************************
LPARAMETER lnRGBColor

lcColor=RIGHT(TRANSFORM(lnRGBColor,"@0"),6)

*** Fox color is BBGGRR, HTML is RRGGBB

RETURN "#" + SUBSTR(lcColor,5,2) + SUBSTR(lcColor,3,2) + LEFT(lcColor,2)
* HTMLColor

************************************************************************
FUNCTION HRef
*********************************
***  Function: Converts a FoxPro Color to an HTML Hex color value
***      Pass: lnRGBColor   -  FoxPro RGB color number - RGB(255,255,255)
***            llNoOutput
***    Return: Hex HTML Color String "#FFFFFF"
************************************************************************
LPARAMETERS lcUrl, lcText, lcAttributes

IF EMPTY(lcText)
   lcText = lcUrl 
ENDIF

IF !EMPTY(lcAttributes)
   RETURN [<a href="] + lcUrl + [" ] + lcAttributes + [>] + lcText + [</a>]
ENDIF

RETURN [<a href="] + lcUrl + [">] + lcText + [</a>]

************************************************************************
FUNCTION MailLink
*****************
***  Function: Returns a safe mail link that is not easily
***            harvested by a spider.
***      Pass:
***    Return:
************************************************************************
LPARAMETERS lcEmail,lcText, lcSubject, lcMessage, lcAttributes, lcImage
LOCAL lcFunc, lnAt, lcEmail1,lcEmail2, lcOutput

lcFunc = SYS(2015)

lnAt = AT("@",lcEmail)
IF lnAt < 1
	RETURN ""
ENDIF	

IF EMPTY(lcAttributes)
	lcAttributes = ""
ENDIF
IF EMPTY(lcImage)
	lcImage = ""
ENDIF	

lcEmail = TRIM(lcEmail)
lcText = TRIM(lcText)
lcEmail1=LEFT(lcEmail,lnAt-1)
lcEmail2=SUBSTR(lcEmail,lnAt+1)

lcOutput = [<a href="javascript:] + lcFunc + [('] + lcEmail1 + [','] + lcEmail2 + ['] 

IF !EMPTY(lcSubject)
	lcOutput = lcOutput +  [,'] +  STRTRAN(UrlEncode(TRIM(lcSubject)),"+"," ")  + [']
ENDIF
IF !EMPTY(lcMessage)
	lcOutput = lcOutput +  [,'] + STRTRAN(UrlEncode(TRIM(lcMessage)),"+"," ") + [']
ENDIF

lcOutput = lcOutput + [);" ] + lcAttributes +[>]

IF !EMPTY(lcImage)
	lcOUtput = lcOutput + lcImage
ELSE
	lcOutput = lcOutput + lcText
ENDIF
		
lcOutput = lcOutput + [</a>] +;
		   [<script> function ] + lcFunc + [(eValue,Text,Subject,Body) { ] + ;
		   [var c1 = 'ma'; var c2 = 'ilto:'; var Link = c1 + c2 + eValue + '@' + Text;] +;
		   [if (Subject) { Link = Link + '?Subject=' + Subject; } if (Body) { Link = Link + '&Body=' + Body; } ] +;
		   [window.open(Link);}</script>]

RETURN lcOutput		   
ENDFUNC
*  wwUtils ::  MailLink

************************************************************************
*  GravatarLink
****************************************
***  Function: Creates an image URL for a given email address
***            that is registered with Gravatar.com.
***
***            Gravatar is a very popular avatar service that
***            requires only an email address to share a picture.
***            Used on many web sites so once you sign up your
***            picture will be used on many sites.
***    Assume: 
***      Pass: lcEmail - Email Address
***            lnSize - Image Size (square) 60-80 is usually good
***            lcDefaultImage - Url to an image if no match is found
***                             for email. Empty shows Gravatar's default
***            lcRating - g, pg, r, x   (Default: pg)
***    Return: URL to the Gravatar image
************************************************************************
FUNCTION GravatarLink(lcEmail,lnSize,lcDefaultImage, lcRating)
LOCAL lcDefaultImage

IF EMPTY(lnSize)
   lnSize = 80
ENDIF

IF !EMPTY(lcEmail)
	lcHash = LOWER(STRCONV(HashMd5(lcEmail),15))
ELSE
  lcHash = ""
ENDIF
IF EMPTY(lcDefaultImage)
   *** Gravatar default image displays
   lcDefaultImage = ""
ELSE
   lcDefaultImage = "&d=" + UrlEncode(lcDefaultImage)
ENDIF   
IF EMPTY(lcRating)
   lcRating = "pg"
ENDIF   

lcUrl = "http://www.gravatar.com/avatar/" + lcHash + "?" +;
       "s=" + TRANSFORM(lnSize) +;
       "&r=" + lcRating +;
       lcDefaultImage

RETURN lcUrl
* GravatarLink


************************************************************************
FUNCTION ContentTypeFromExtension
*********************************
***  Function: Converts an extension to a Content Type
***      Pass: lcExtension = Extension without the dot
***    Return: Content Type string
************************************************************************
LPARAMETERS lcExtension
LOCAL lcContentType

lcExtension = LOWER(lcExtension)

DO CASE
  * Text types:
  CASE LIKE("*htm*", lcExtension)
    lcContentType = "text/html"
  CASE lcExtension == "xml" OR lcExtension == "rtf" OR lcExtension == "css" 
    lcContentType = "text/" + lcExtension
  * Application types:
  CASE lcExtension == "pdf"
    lcContentType = "application/pdf"
  CASE lcExtension == "fdf"
    lcContentType = "application/vnd.fdf"
  CASE lcExtension == "doc" OR lcExtension == "dot"
    lcContentType = "application/msword"
  CASE lcExtension == "ppt"
    lcContentType = "application/vnd.ms-powerpoint"
  CASE lcExtension == "xls" OR lcExtension == "iqy"
    lcContentType = "application/vnd.ms-excel"
  CASE lcExtension == "txt" OR lcExtension == "asc" OR lcExtension == "sdf" OR lcExtension == "prg"
    lcContentType = "text/plain"
  CASE lcExtension == "file" OR lcExtension == "exe" OR lcExtension == "dll" OR lcExtension == "bin"
    lcContentType = "application/octet-stream"
  CASE LIKE("wp*", lcExtension)
    lcContentType = "application/wordperfect5.1"
  CASE LIKE("vs?", lcExtension)
    lcContentType = "application/vnd.visio"
  CASE lcExtension == "js"
    lcContentType = "application/x-javascript"
  CASE lcExtension == "json"
    lcContentType = "application/json"
  CASE lcExtension == "zip"
    lcContentType = "application/zip"
  * Image types:
  CASE LIKE("jp*g", lcExtension)
    lcContentType = "image/jpeg"
  CASE LIKE("tif*", lcExtension)
    lcContentType = "image/tiff"
  CASE lcExtension == "gif" OR lcExtension == "bmp" OR lcExtension == "tng" 
    lcContentType = "image/" + lcExtension
  CASE LIKE("svg*", lcExtension)
    lcContentType = "image/svg-xml"
  * Default?:
  OTHERWISE
	loAPi = CREATEOBJECT("wwAPI")
	lcContentType = loAPI.ReadRegistryString(HKEY_CLASSES_ROOT,"." + lcExtension,"Content Type")
	IF ISNULL(lcContentType)
	  lcContentType = "application/octet-stream"
	ENDIF
ENDCASE

RETURN lcContentType  



************************************************************************
FUNCTION CharToBin
******************
***  Function: Converts a DWORD value in binary string form back into
***            a numeric value
***      Pass: tcWord  -    Binary string value (from a structure?)
***    Return: numeric value of binary string
*************************************************************************
LPARAMETER tcWord

  LOCAL i, lnWord

  lnWord = 0
  FOR i = 1 TO LEN(tcWord)
    lnWord = lnWord + (ASC(SUBSTR(tcWord, i, 1)) * (2 ^ (8 * (i - 1))))
  ENDFOR

RETURN lnWord

************************************************************************
PROCEDURE StrTranC
******************
***  Function: Like Strtran but case insensitive
***      Pass: lcString   -  Entire string
***			   lcDelim1   -  String to replace
***            lcDelim2	  -  String to replace with
***    Return: translated string
*************************************************************************
LPARAMETER lcString, lcSource, lcReplace

#IF wwVFPVersion > 6
  RETURN STRTRAN(lcString,lcSource,lcReplace,1,-1,1)
#ELSE
LOCAL lnAt

lnAt = 1
lnReplaceSize = LEN(lcSource)
DO while .T.
   lnAt = ATC(lcSource,lcString)
   IF lnAT = 0
      RETURN lcString
   ENDIF
   
   lcString = STUFF(lcString,lnAt,lnReplaceSize,lcReplace)
ENDDO

RETURN lcString
#ENDIF

************************************************************************
FUNCTION TimeToCStrict
****************************************
***  Function: Returns time in strict time format or in SQL Server
***            compliant format (llSql flag set).
***      Pass: Datetime value
***    Return: strict DateTime string, or SQL Server time string
************************************************************************
LPARAMETER ltTime, llSQL

IF llSQL
  IF EMPTY(ltTime)
     RETURN "''"
  ENDIF
  
  RETURN "'" + TRANSFORM(TTOC(ltTime,1),"@R 9999-99-99 99:99:99") + "'"
ENDIF

IF EMPTY(ltTime)
   RETURN "{^}"
ENDIF

RETURN "{^" + TRANSFORM(TTOC(ltTime,1),"@R 9999-99-99 99:99:99") + "}"

*!*   lcCentury = SET("CENTURY")
*!*   lcDateMode = SET("DATE")
*!*   SET CENTURY ON
*!*   SET DATE TO YMD

*!*   IF llSQL
*!*     lcDate = "'" + TRANSFORM(TTOC(ltTime,1),"@R 9999-99-99 99:99:99") + "'"
*!*   ELSE
*!*     lcDate = "{^" + TTOC(ltTime) + "}"
*!*   ENDIF

*!*   SET DATE TO &lcDateMode
*!*   SET CENTURY &lcCentury
*!*   RETURN lcDate

************************************************************************
FUNCTION DateToC
******************
***  Function: Converts a date to string displaying empty dates as blanks
***            rather than displaying the empty date format
***      Pass: ldDate  - Date to display
***    Return: Date String or "" if invalid date
*************************************************************************
LPARAMETER ldDate

IF EMPTY(ldDate)
   RETURN ""
ENDIF

RETURN  DTOC(ldDate)
* DateTOC



************************************************************************
FUNCTION TimeToC
******************
***  Function: Converts a time to string displaying empty  as blanks
***            and formatting the time string properly
***            NOTE: Assumes US time format with AM/PM extension
***      Pass: ltTime  - Date to display (Pass Time, Date or Char)
***    Return: Time String or "" if invalid date (Year is not returned)
*************************************************************************
LPARAMETER ltTime
LOCAL lcTimeStamp, lnHour, lcAm, lcType

IF EMPTY(ltTime)
   RETURN ""
ENDIF

lcType = VARTYPE(ltTime)

IF lcType $ "DT"
  lcTimestamp = TTOC(ltTime)
  lcTimeStamp = StringFormat("{0}. {1}, {2}",LEFT(CMONTH(ltTime),3),DAY(ltTime),YEAR(ltTime))
ELSE
  RETURN ltTime
ENDIF  

IF lcType = "D"
   RETURN lcTimeStamp
ENDIF

IF SET("DATE") = "AMERICAN"
	lnHour = HOUR(ltTime)
	lcAm = "am"
	IF lnHour > 12
	   lnHour = lnHour - 12
	   lcAm = "pm"
	ENDIF
	IF lnHour = 0
      lnHour = 12
      lcAm = "am"
   ENDIF   
   lcTimeStamp = lcTimeStamp + StringFormat(" {0}:{1}{2}",HOUR(ltTime),PADL(MINUTE(ltTime),2,"0"),lcAm)
ELSE
	lcTimeStamp = lcTimeStamp + StringFormat(" {0}:{1}",HOUR(ltTime),PADL(MINUTE(ltTime),2,"0"))
ENDIF

RETURN lcTimeStamp
* TimeToC

************************************************************************
*  ShortTime
****************************************
***  Function: Creates a short date string in the format of 
*** 		   01:10pm
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ShortTime(ltTime)
LOCAL lcTime

IF EMPTY(ltTime)
   RETURN ""
ENDIF
   
lcTime = TTOC(ltTime,2)

RETURN  LOWER(STUFF(lcTime,6,3,""))
ENDFUNC
*   ShortTime


************************************************************************
*  ShortDate
****************************************
***  Function:
***    Assume:
***      Pass:  lnMode = 0 - Just month and Day, 1 - Month Day Year, 2 - Month day year and time
***    Return:
************************************************************************
FUNCTION ShortDate(ltDate,lnMode)
LOCAL lcTime
 
IF EMPTY(lnMode)
   lnMode = 0
ENDIF
IF EMPTY(ltDate)
  RETURN ""
ENDIF  

lcDate = LEFT(CMONTH(ltDate),3) + ". " + TRANSFORM(DAY(ltDate)) 

*** Return short date and short time
IF lnMode == 4
	lcDate = lcDate + " - " + ShortTime(ltDate)
	RETURN lcDate
ENDIF


*** Add Year
IF lnMode > 0
	lcDate = lcDate + ", " + TRANSFORM(YEAR(ltDate))
ENDIF	

*** Add Time
IF lnMode > 1
  lcDate = lcDate + "  " + ShortTime(ltDate)
ENDIF


RETURN lcDate
ENDFUNC
*   ShortDate

************************************************************************
*  NoEmptyDate
****************************************
***  Function: Returns a replacement datetime value for an
***            empty date. use for XML conversions or when
***            returning data for SQL server - anything that
***            doesn't support empty dates.
***    Assume:
***      Pass: 
***    Return:
************************************************************************
FUNCTION NoEmptyDate(ltDateTimeValue, ltDateTime)

IF ISNULL(ltDateTimeValue) OR !EMPTY(ltDateTimeValue) 
   RETURN ltDateTimeValue
ENDIF   

IF EMPTY(ltDateTime)
   ltDateTime = {^1900-01-01}
ENDIF

RETURN ltDateTime
ENDFUNC
*   NoEmptyDate

************************************************************************
* wwUtils :: MimeDateTime
****************************************
***  Function: Converts Mime Dates to and from VFP dates
***    Assume:
***      Pass: Mime Date or VFP DateTime value
***    Return:
************************************************************************
FUNCTION MimeDateTime(lvDateTime,llGMT)
LOCAL lcDays, lcMonths, ;
   lnCount, ;
   lcMonth, ;
   ltTime, ;
   lnOffset, ;
   lnOff, ;
   lnDiff

lcMonths = "JanFebMarAprMayJunJulAugSepOctNovDec"
lcDays = "SunMonTueWedThuFriSat"

IF VARTYPE(lvDateTime) = "C"
   LOCAL ltResult as DateTime
   TRY	
	   *** Convert a MimeDate String to a DateTime value
	   DIMENSION laParts[1]
	   lnCount = APARSESTRING(@laParts,lvDateTime," ")
	   IF lnCount < 5
	      RETURN {}
	   ENDIF

	   IF ATC(STRTRAN(RTRIM(laParts[1]),",",""),lcDays) < 1
	   		*** Resize and insert a row at the first location
			DIMENSION laParts[ALEN(laParts,1) + 1]
	   		AINS(laParts,1)
	   ENDIF
	   
	   lcMonth  = TRANSFORM((ATC(laParts[3],lcMonths)+2) / 3 )

	   *ltTime = CTOT( lcMonth + "/" + laParts[2] + "/" + laParts[4] + " " + laparts[5] )
	   ltTime = DATETIME(VAL(laParts[4]),VAL(lcMonth), VAL(laParts[2]),;
	                     VAL(LEFT(laParts[5],2)),VAL(SUBSTR(laParts[5],4,2)), VAL(RIGHT(laParts[5],2)))

	   IF lnCount < 6
	      *** No time Date Offset
	      RETURN ltTime
	   ENDIF
	   lnOffset = GetTimeZone() / 60
	   lnOff = VAL(laParts[6]) /100
	   lnDiff = (lnOffset + lnOff) * 3600
	   
	   ltResult = ltTime - lnDiff
   CATCH
   	   *** Invalid date
   	   ltResult = { : }
   ENDTRY
   RETURN ltResult
ELSE
   IF VARTYPE(lvDateTime) = "D"
    	lvDateTime = DTOT(lvDateTime)
   ENDIF

   lnOffset = GetTimeZone()/60

   IF EMPTY(lvDateTime) 
      lvDateTime = DATETIME()
   ENDIF
   
   IF llGMT
      *** Take off the Timeoffset
      lvDateTime = lvDateTime + (lnOffSet * 3600)
   ENDIF      
   
   *** Create a MIME Date from a VFP date
   lnDay = DOW(lvDateTime)
   IF lnDay = 0
      RETURN ""
   ENDIF

   lcDays = "SunMonTueWedThuFriSat"
   lcMime = SUBSTR(lcDays,( (lnDay-1) * 3)+1,3) + ", " + ;
            TRANSFORM(DAY(lvDateTime)) + " "

   lcMime = lcMime + SUBSTR(lcMonths,( ( MONTH(lvDateTime)-1) * 3) + 1,3) + " " +; 
            TRANSFORM(YEAR(lvDateTime)) + " " + ;
            PADL(HOUR(lvDateTime),2,"0") + ":" + PADL(MINUTE(lvDateTime),2,"0") + ":" +;
            PADL(SEC(lvDateTime),2,"0")

   
   IF llGMT
      RETURN lcMime + " GMT"
   ELSE
      RETURN lcMime + IIF(lnOffset>0," -"," +") + PADL(TRANSFORM(ABS(lnOffSet)),2,"0") + "00"
   ENDIF
ENDIF

ENDFUNC

************************************************************************
FUNCTION GetAppStartPath
*********************************
***  Function: Returns the FoxPro start path
***            of the *APPLICATION*
***            under all startmodes supported by VFP.
***            Returns the path of the starting EXE, 
***            DLL, APP, PRG/FXP
***    Return: Path as a string with trailing "\"
************************************************************************

DO CASE 
   *** VFP 6 and later provides ServerName property for COM servers EXE/DLL/MTDLL
   CASE INLIST(_VFP.StartMode,2,3,5) 
         lcPath = JustPath(_VFP.ServerName)

   *** Interactive
   CASE (_VFP.StartMode) = 0
         lcPath = SYS(5) + CURDIR()
         
   *** Active Document
   CASE ATC(".APP",SYS(16,0)) > 0
       lcPath = JustPath(SYS(16,0))

  *** Standalone EXE or VFP Development
  OTHERWISE
       lcPath = JustPath(SYS(16,0))
       IF ATC("PROCEDURE",lcPath) > 0
         lcPath = SUBSTR(lcPath,RAT(":",lcPath)-1)
       ENDIF
ENDCASE

RETURN AddBs(lcPath)
* EOF GetAppStartPath      


************************************************************************
FUNCTION ShortPath
******************
***  Function: Converts a Long Windows filename into a short
***            8.3 compliant path/filename
***      Pass: lcPath   -  Path to check
***    Return: lcShortFileName
*************************************************************************
LPARAMETER lcPath

DECLARE INTEGER GetShortPathName IN Win32API;
  STRING @lpszLongPath,	STRING @lpszShortPath,;
  INTEGER cchBuffer

lcPath = lcPath
lcshortname = SPACE(260)
lnlength = LEN(lcshortname)
lnresult = GetShortPathName(@lcPath, @lcshortname, lnlength)
IF lnResult = 0
   RETURN ""
ENDIF
RETURN LEFT(lcShortName,lnResult)

************************************************************************
FUNCTION LongPath
****************************************
***  Function: Returns the Long path name of a short path and properly
***            cases a path string based on directory entries.
***            Works off full or relative paths.
***      Pass: lcFilename - name of the file or path to case
***    Return: Properly cased path, or if file doesn't exist 
***            same string that was input.
************************************************************************
LPARAMETERS lcFileName
LOCAL lcFile, lcShortFileName, lnResult

DECLARE INTEGER GetLongPathName IN WIN32API  ;
  STRING lpFileName, ;
  STRING @lcBuffer, ;
  INTEGER lnBufferSize
  
lcFile=SPACE(255)

*** Must convert to ShortPath first in order for GetLongPathName to work
*** properly on filenames longer than 8 characters
lcShortFilename = ShortPath(lcFilename)
IF EMPTY(lcShortFileName)
  RETURN lcFileName
ENDIF  

lnResult = GetLongPathName(lcShortFileName,@lcFile,255)
IF lnResult = 0
   RETURN lcFileName
ENDIF
   
RETURN LEFT(lcFile,lnResult)
* LongPath()
   

************************************************************************
FUNCTION DeleteFiles
********************
***  Function: Returns the size of a file
***      Pass: lcFileName   - Wildcard File Spec  (d:\temp\*.pdf)
***            lnTimeout    - Timeout in seconds 
***    Return: the size of the file or -1 on error
************************************************************************
PARAMETERS lcFileSpec, lnTimeout
LOCAL lnX,lnFiles, loAPI, loEval

lnTimeout=IIF(EMPTY(lnTimeout),300,lnTimeout)

lnFiles = aDir(laFiles,lcFileSpec)

DECLARE INTEGER DeleteFile IN WIN32API ;
    STRING

FOR lnX=1 to lnFiles
*   ldtime = CTOT( DTOC(laFiles[lnX,3]) + " " + laFiles[lnX,4] )
    ldTime = DATETIME(YEAR(laFiles[lnX,3]), MONTH(laFiles[lnx,3]), DAY(laFiles[lnX,3]), ;
                      VAL(LEFT(laFiles[lnX,4],2)),;
                      VAL(SUBSTR(laFiles[lnX,4],4,2)), ;
                      VAL( RIGHT(laFiles[lnX,4],2)))

   IF  ldTime + lnTimeout < DateTime()
     DeleteFile( ADDBS(justpath(lcFileSpec)) + laFiles[lnX,1] )
     *ERASE (ADDBS(justpath(lcFileSpec)) + laFiles[lnX,1]) 
   ENDIF
ENDFOR

RETURN .T.
ENDFUNC


************************************************************************
FUNCTION IsDir
******************
***  Modified: 10/09/97
***  Function: Checks to see whether a directory exists
***      Pass: lcPath   -  Path to check
***    Return: .T. or .F.
*************************************************************************
LPARAMETER lcPath
DIMENSION laTemp[1]
IF ADIR(laTemp,lcPath,"DH") < 1
   RETURN .F.
ENDIF
RETURN .T.

************************************************************************
FUNCTION FileTime(lcFileName)
****************************************
***  Function: Returns the file datetime for a specified file
***    Assume:
***      Pass: Filename 
***    Return: Date Time for the the file's last update or empty date
************************************************************************
LOCAL lnFile, ldDate
LOCAL ARRAY laFile[1]
lnFile = ADIR(laFile,lcFileName)
IF lnFile = 0
   RETURN {:}
ENDIF   
ldDate = laFile[1,3]
RETURN EVALUATE("{^" + TRANSFORM(YEAR(ldDate)) + "-" + ;
               TRANSFORM(MONTH(ldDate)) + "-" + ;
               TRANSFORM(DAY(ldDate)) + " " + laFile[1,4] + "}")
*   FileTime

************************************************************************
FUNCTION FileAttributes(lcFileName)
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
LOCAL lnFile, ldDate
LOCAL ARRAY laFile[1]
lnFile = ADIR(laFile,lcFileName)
IF lnFile = 0
   RETURN ""
ENDIF   
RETURN laFile[1,5]
ENDFUNC
*   FileAttributes

************************************************************************
FUNCTION FileSize
******************
***  Function: Returns the size of a file
***      Pass: lcFileName
***    Return: the size of the file or -1 on error
************************************************************************
LPARAMETERS lcFileName
LOCAL lh, lnSize

lh = FOPEN(lcFileName,0)
IF lh = -1
   RETURN -1
ENDIF

lnSize = FSEEK(lh, 0, 2)   

FCLOSE(lh)

RETURN lnSize
*EOP FileSize


************************************************************************
FUNCTION Slash
******************
***  Function: Converts slashes from DOS -> Web and vice versa
***      Pass: lcPath   -  Path to convert
***            lcStyle  -  "WEB" or "DOS"
***    Return: update path
************************************************************************
LPARAMETER lcPath, lcStyle
lcStyle=IIF(type("lcStyle")="C",UPPER(lcStyle),"")
IF lcStyle="WEB"
   lcPath=CHRTRAN(lcPath,"\","/")
ELSE
   lcPath=CHRTRAN(lcPath,"/","\")
ENDIF
RETURN lcPath
*EOP LPARAMETER



************************************************************************
FUNCTION ProgLevel
******************
***  Function: Returns the current Calling Stack level. Used to check
***            recursive Error calls in Error methods.
*************************************************************************

FOR lnX=1 to 128
   IF EMPTY(SYS(16,lnX))
      exit
   ENDIF
ENDFOR && lnX=1 to 128

*** -1 for lnX count - -1 for ProgLevel Call
RETURN lnX - 2


************************************************************************
FUNCTION AParseString
**********************
***  Modified: 07/03/97
***  Function: Parses a delimited string into an array
***      Pass: laResult    -   Array containing the result strings (@)
***            lcString    -   The full string
***            lcDelimiter -   The delimiter string
***    Return: Count of strings or 0 if null string is passed
*************************************************************************
LPARAMETER laResult, lcString, lcDelimiter
LOCAL lnLastPos, lnItemCount, i
lnItemCount = ALINES(laResult,lcString,lcDelimiter)
RETURN lnItemCount

************************************************************************
FUNCTION URLDecode
******************
***  Function: URLDecodes a text string to normal text.
***    Assume: Uses wwIPStuff.dll
***      Pass: lcText       -   Text string to decode
***            llUtf8Decode -   Decodes from UTF-8
***    Return: Decoded string or ""
************************************************************************
LPARAMETERS lcText,llUtf8Decode
LOCAL lnSize, lnLoc, lcHex, lnHex, lcRetval, loEval, loRegEx

loRegEx = GetRegExObject()

*** Use wwIPStuff for large buffers
IF LEN(lcText) > 255
   DECLARE INTEGER URLDecode ;
      IN WWIPSTUFF AS API_URLDecode ;
      STRING @cText

   lnSize=API_URLDecode(@lcText)

   IF lnSize > 0
      lcText = SUBSTR(lcText,1,lnSize)
   ELSE
      lcText = ""
   ENDIF

   IF llUtf8Decode
   	  RETURN STRCONV(lcText,11)	
   ENDIF
   
   RETURN lcText
ENDIF


*** First convert + to spaces
lcText=STRTRAN(lcText,"+"," ")

*** Handle Hex Encoded Control chars
lcRetval = ""
lnLoc = 1
DO WHILE .T.
   *** Format: %0A  ( CHR(10) )
   lnLoc = AT('%',lcText)

   *** No Hex chars - just return the text
   IF lnLoc > LEN(lcText) - 2 OR lnLoc < 1
      lcRetval = lcRetval + lcText
      EXIT
   ENDIF

   *** Now read the next 2 characters
   *** Check for digits - at this point we must have hex pair!
   lcHex=SUBSTR(lcText,lnLoc+1,2)
   
   lnHex = -1
   TRY 
		lnHex = EVALUATE("0x" + lcHex)
   CATCH
   ENDTRY
   
   
   *** Now concat the string plus the evaled hex code
   IF lnHex > -1
      lcRetval = lcRetval + LEFT(lcText,lnLoc-1) + ;
         CHR( EVAL("0x"+lcHex) )

      *** Trim out the input string
      IF LEN(lcText) > lnLoc + 2
         lcText = SUBSTR(lcText,lnLoc+3)
      ELSE
         EXIT
      ENDIF
   ELSE
      lcRetVal = lcRetVal + SUBSTR(lcText,1,lnLoc)
      lcText = SUBSTR(lcText,lnLoc+1)  && Skip over
   ENDIF
ENDDO

IF llUtf8Decode
	RETURN STRCONV(lcRetval,11)
ENDIF

RETURN lcRetval
ENDFUNC
* EOF URLDecode

************************************************************************
* wwUtils :: HtmlDecode
****************************************
***  Function: Decodes HTML Encoding like &lt;, &gt;, &amp;,&quote;
***            and &#0A hex values
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION HtmlDecode(lcText)

lcText = STRTRAN(lcText,"&lt;","<")
lcText = STRTRAN(lcText,"&gt;",">")
lcText = STRTRAN(lcText,"&nbsp;"," ")
lcText = STRTRAN(lcText,"&quot;",["])
lcText = STRTRAN(lcText,"&amp;",[&])

lcRetval = ""
DO WHILE .T.
   *** Format: &#0A;  ( CHR(10) )
   lnLoc = AT('&#',lcText)

   *** No Hex chars - just return the text
   IF lnLoc > LEN(lcText) - 2 OR lnLoc < 1
      lcRetval = lcRetval + lcText
      EXIT
   ENDIF

   *** Now read the next 2 characters
   *** Check for digits - at this point we must have hex pair!
   lcHex=SUBSTR(lcText,lnLoc+2,2)

   *** Now concat the string plus the evaled hex code
   lcRetval = lcRetval + LEFT(lcText,lnLoc-1) + ;
      CHR( EVAL(lcHex) )

   *** Trim out the input string
   IF LEN(lcText) > lnLoc + 2
      lcText = SUBSTR(lcText,lnLoc+5)
   ELSE
      EXIT
   ENDIF
ENDDO

RETURN lcRetVal
ENDFUNC
*  wwUtils :: HtmlDecode

************************************************************************
FUNCTION GetURLEncodedKey
*********************************
***  Function: Retrieves a 'parameter' from the query string that
***            is encoded with standard CGI/ISAPI URL encoding.
***            Typical URL encoding looks like this:
***
***    "User=Rick+Strahl&ID=0011&Address=400+Morton%0A%0DHood+River"
***
***      Pass: lcVal   -   Form Variable to retrieve
***    Return: Value or ""
************************************************************************
LPARAMETERS tcURLString, lcKey, llNullIfNotFound
LOCAL lnLoc,c2, cStr, lcURLString, lcRetval, llQ

lcURLString=IIF(EMPTY(tcURLString),"","&"+tcURLString)
IF EMPTY(lcKey)
   RETURN ""
ENDIF
lcKey=STRTRAN(lcKey," ","+")

llQ = .f.
lnAt = ATC("&" + lcKey + "=",tcUrlString) 
IF lnAt < 1 
   lnAt = ATC("?" + lcKey + "=",tcUrlString) 
   llQ = .T.
ENDIF   

IF lnAt < 1 AND llNullIfNotFound
	RETURN NULL
ENDIF	

IF !llQ
	lcRetval = STREXTRACT(lcUrlString,"&"+lcKey+"=","&",1,3)
ELSE	
    lcRetval = STREXTRACT(lcUrlString,"?"+lcKey+"=","&",1,3)
ENDIF	

IF EMPTY(lcRetVal)
   RETURN ""
ENDIF
   	
RETURN URLDecode(lcRetval)
ENDFUNC


************************************************************************
* wwNameValueCollection ::  GetUrlEncodedValues
****************************************
***  Function: Returns a collection of all keys and values
***            in a string of urlencoded key values.
***            Returns a wwNameValueCollection.
***    Assume: Useful for parsing querystrings or Form data
***            into an easily parsable collection you can walk through
***      Pass: lcVars    - String of UrlEncoded values
***            lcPrefix  - Optional prefix for variables to retrieve
***    Return: wwNameValueCollection
************************************************************************
FUNCTION GetUrlEncodedValues(lcVars,lcPrefix)
LOCAL loVars, lnX, lnAt, lcPoint,lnEqual,lcKey, lcValue

loVars = CREATEOBJECT("wwNameValueCollection")
  
IF EMPTY(lcPrefix)
	lcPrefix = ""
ENDIF
  
lcPointer = "&"+lcVars
  
lnAt = ATC("&" + lcPrefix,lcPointer)
IF lnAt = 0
   RETURN 0
ENDIF

lcPointer = SUBSTR(lcPointer,lnAt+1)

DO WHILE lnAt > 0
	*** Find = sign then extract the value
	lnEqual = AT("=",lcPointer)

	*** No Equal Sign - Invalid key so skip it
	IF lnEqual = 0
	   lnAt=ATC("&" + lcPrefix,lcPointer)
	   LOOP
	ENDIF
	   
	lcKey = SUBSTR(lcPointer,1,lnEqual-1)

	lcPointer = SUBSTR(lcPointer,lnEqual+1)

	*** Find the & at the end
	lnAt = ATC("&" + lcPrefix,lcPointer)
	IF lnAt = 0
	   lcValue = lcPointer
	ELSE
	   lcValue = LEFT(lcPointer + lcPrefix,lnAt - 1)
	ENDIF

	lcValue = URLDecode(lcValue) 
	lcKey = URLDecode(lcKey)

	loVars.Add(lcKey,lcValue)
	  
	lcPointer = SUBSTR(lcPointer,lnAt + 1 )
ENDDO  

RETURN loVars  
ENDFUNC
*  wwNameValueCollection ::  GetUrlEncodedValues


************************************************************************
*  RemoveUrlEncodedKey
****************************************
***  Function: Removes a single URL encoded key from a string
***      Pass: lcQuery  -  QUery string  ?
***    Return:
************************************************************************
FUNCTION RemoveUrlEncodedKey(lcQuery,lcKey)
LOCAL lcVal, llStartAdded,lcLeft

IF EMPTY(lcQuery)
   RETURN lcQuery
ENDIF

llStartAdded = .F.
lcLeft = LEFT(lcQuery,1)
IF !INLIST(lcLeft,"?","&")
   lcQuery = "&" + lcQuery
   llStartAdded = .T.
ENDIF   
   
*** Get the value
lcVal = STREXTRACT(lcQuery,"&" + lcKey + "=","&",1,3)
IF EMPTY(lcVal)
   *** Select whole string and delims
   lcVal = STREXTRACT(lcQuery,"?" + lcKey + "=","&",1,7)
   
   IF !EMPTY(lcVal)
	*** Replace ? and trailing & along with content\
	lcQuery = STRTRAN(lcQuery,lcVal,"?")
   ENDIF
ELSE
	*** Replace only the key   
	lcQuery = STRTRAN(lcQuery,"&" + lcKey + "=" + lcVal)
ENDIF


IF llStartAdded AND LEFT(lcQuery,1) = "&"
  RETURN SUBSTR(lcQUery,2)
ENDIF

RETURN lcQuery
ENDFUNC
*   RemoveUrlEncodedKey

************************************************************************
*  SetUrlEncodedKey
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION SetUrlEncodedKey(lcQuery, lcKey, lcValue)
LOCAL lcFullValue, lnLoc, lnLeft, lcLeft, lcRight, llStarted

lcFullValue = ""

llStartAdded = .F.
lcLeft = LEFT(lcQuery,1)
IF !INLIST(lcLeft,"?","&")
   lcQuery = "&" + lcQuery
   llStartAdded = .T.
   lcLeft = "&"
ENDIF

lcValue = URLEncode(lcValue)

lnLoc=ATC("&"+lcKey+"=",lcQuery)
IF lnLOc < 1
   lnLoc = ATC("?"+lcKey+"=",lcQuery)
   IF lnLoc > 0
	   lcLeft = "?"
   ENDIF   
ELSE
   lnLeft = "&"   
ENDIF   

if lnLoc > 0
   lcRest = SUBSTR(lcQuery,lnLoc)
   lnLength = ATC("&",lcRest,2) 
   IF lnLength=< 1 
      *** No & at end - full string size
      lnLength = LEN(lcRest) + 1 && One to long to match &
   ENDIF
   lcFullValue = LEFT(lcRest,lnLength-1)
   
   IF llStartAdded
   		lcQuery = SUBSTR(lcQuery,2)
   ENDIF
   
   RETURN STRTRAN(lcQuery,lcFullValue,lcLeft+lcKey+"="+lcValue)
ENDIF

lcRight = RIGHT(lcQUery,1)
DO CASE
	CASE EMPTY(lcQuery)
	   lcQuery = "?"	
	CASE lcRight = "?" OR lcRight = "&"
	   lcLeft = ""
	OTHERWISE
	   lcLeft = "&"
ENDCASE

IF llStartAdded
   lcQuery = SUBSTR(lcQuery,2)
ENDIF   

RETURN lcQuery + lcLeft + lcKey +"="+lcValue
ENDFUNC
* EOF wwRequest::SetUrlEncodedKey


********************************************************
FUNCTION URLEncode
*******************
***  Function: Encodes a string in URL encoded format
***            for use on URL strings or when passing a
***            POST buffer to wwIPStuff::HTTPGetEx
***      Pass: tcValue  -   String to encode
***    Return: URLEncoded string or ""
********************************************************
LPARAMETER tcValue
LOCAL lcResult, lcChar, lnSize, lnX

*** Large Buffers use the wwIPStuff function 
*** for quicker response
if  LEN(tcValue) > 512
   lnSize=LEN(tcValue)
   tcValue=PADR(tcValue,lnSize * 3)

   DECLARE INTEGER VFPURLEncode ;
      IN WWIPSTUFF ;
      STRING @cText,;
      INTEGER cInputTextSize
   
   lnSize=VFPUrlEncode(@tcValue,lnSize)
   
   IF lnSize > 0
      RETURN SUBSTR(TRIM(tcValue),1,lnSize)
   ENDIF
   RETURN ""
ENDIF   
   
*** Do it in VFP Code
lcResult=""

FOR lnX=1 to len(tcValue)
   lcChar = SUBSTR(tcValue,lnX,1)
   DO CASE
   CASE ATC(lcChar,"ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-._~") > 0
      lcResult=lcResult + lcChar
   CASE lcChar=" "
      lcResult = lcResult + "%20"
   OTHERWISE
	   *** Convert others to Hex equivalents
	   lcResult = lcResult + "%" + RIGHT(transform(ASC(lcChar),"@0"),2)
   ENDCASE
ENDFOR && lnX=1 to len(tcValue)

RETURN lcResult
* EOF URLEncode

************************************************************************
*  SafeEval
****************************************
***  Function: A Safe Evaluation function that wraps the EVAL()
***            into a TRY/CATCH so no outward failure occurs
***    Assume:
***      Pass: lcEvalString  -  VFP Expression to evaluate
***            @lcErrorResult -  Determines what's returned on error
***                             Can be ERRORMESSAGE to return error
***                             message string
***    Return: evaluated result or error
************************************************************************
FUNCTION SafeEval(lcEvalString,lvErrorResult)
LOCAL lvResult, loException

TRY
	lvResult = EVALUATE(lcEvalString)
CATCH TO loException
	DO CASE
	CASE VARTYPE(lvErrorResult) = "C" AND ;
	     lvErrorResult = "ERRORMESSAGE"
	   lvResult = loException.Message
	OTHERWISE
	   lvResult = lvErrorResult
	ENDCASE
ENDTRY

RETURN lvResult
*   SafeEval

************************************************************************
*  SafeCommand
****************************************
***  Function: Excecutes a single VFP command wrapped in the context
***            of a TRY/CATCH so it won't fail/cause an error
***    Assume:
***      Pass:  lcCommand - single VFP Command or expression
***             lcErrorResult - optional result to return on error
***    Return: "" or error message
************************************************************************
FUNCTION SafeCommand(lcCommand,lcErrorResult)
LOCAL lcResult, loException

lcResult = ""
TRY
	&lcCommand
CATCH TO loException
    IF EMPTY(lcErrorResult)
       lcErrorResult = ""
    ENDIF
	DO CASE
	CASE lcErrorResult = "ERRORMESSAGE"
	   lcResult = "Error: " + loException.Message
	OTHERWISE
	   lcResult = lcErrorResult
	ENDCASE
ENDTRY

RETURN lcResult
ENDFUNC
*   SafeCommand

************************************************************************
FUNCTION MergeText
******************
***  Function: This function provides an evaluation engine for FoxPro
***            expressions and Codeblocks that is tuned for Active
***            Server syntax. It works with any delimiters however. This
***            parsing engine is faster than TEXTMERGE and provides
***            extensive error checking and the ability to run
***            dynamically in the VFP runtime (ie uncompiled). Embed any
***            valid FoxPro expressions using
***            
***               <%= Expression %>
***            
***            and any FoxPro code with
***            
***               <% CodeBlock %>
***            
***            Expressions ideally should be character for optimal
***            speed, but other values are converted to string
***            automatically. Although optimized for the delimiters
***            above you may specify your own. Make sure to set the
***            llNoAspSyntax parameter to .t. to disallow the = check
***            for expressions vs code. If you use your own parameters
***            you can only evaluate expressions OR you must use ASP
***            syntax and follow your deleimiters with an = for
***            expressions.
***   Assume:  Delimiter is not used in regular text of the text.
***     Uses:  wwEval class (wwEval.prg)
***            Codeblock Class (wwEval.prg)         
***     Pass:  tcString    -    String to Merge
***            tcDelimeter -    Delimiter used to embed expressions
***                             Default is "<%"
***            tcDelimeter2-    Delimiter used to embed expressions
***                             Default is "%>"
***            llNoAspSytnax    Don't interpret = following first
***                             parm as expression. Everything is 
***                             evaluated as expressions.
***
***  Example:  loHTML.MergeText(HTMLDocs.MemField,"##","##",.T.)
*************************************************************************
LPARAMETER tcString,tcDelimiter, tcDelimiter2, llNoASPSyntax
IF VARTYPE(__goEval) != "O"
	PUBLIC __goEval
	__goEval = CREATE([WWC_wwEval])
ENDIF	
RETURN __goEval.MergeText(@tcString,tcDelimiter, tcDelimiter2, llNoASPSyntax)
*EOF MergeText

************************************************************************
*  MergeTextFromFile
****************************************
***  Function: Merges ASP style syntax from a file on disk.
***    Assume: see MergeText for more info
***      Pass: 
***    Return:
************************************************************************
FUNCTION MergeTextFromFile
LPARAMETERS tcFile, tlCacheTemplate, tcDelimiter, tcDelimiter2, llNoAspSyntax
IF tlCacheTemplate
	RETURN MergeText(CacheFile(tcFile,0),tcDelimiter,tcDelimiter2,llNoAspSyntax)
ENDIF	
RETURN MergeText(File2Var(tcFile),tcDelimiter,tcDelimiter2,llNoAspSyntax)
ENDFUNC
*   MergeTextFromFile




************************************************************************
* wwUtils ::  RenderAspScript
****************************************
***  Function: Renders an ASP style template or string
***    Assume:
***      Pass: lcTemplate - Name of hte template to render
***                         OR string of the templ if llIsString is .T.
***            lnMode     - 1 - Dynamic
***                         2 - Precompiled
***                         3 - ExecScript
***            llIsString  - determines if lcTemplate is a string or file
***    Return: output from template. On error an HTML error page 
***            is returned
************************************************************************
FUNCTION RenderAspScript(lcTemplate,lnMode,llIsString,llUseHttpResponse)
LOCAL loScript

IF EMPTY(lnMode) 
   lnMode = 1
ENDIF
IF llIsString
   *** String has to execute with ExecuteScript
   lnMode = 3
ENDIF

loScript = CREATEOBJECT("wwScripting")
IF llUseHttpResponse
   *** include HTTP header in the response
   loScript.cScriptResponseClass = "wwScriptingHttpResponse"  
ENDIF

DO CASE
    *** Create PRG and FXP files if version has changed
 	CASE lnMode = 1
		RETURN loScript.RenderAspScript( lcTemplate )

    *** Execute only the FXP file - no version check no file check
    CASE lnMode = 2
    	 *** Just run the script code without compiling or checking	
         loScript.lNoVersionCheck = .T.
         RETURN loScript.RenderAspScript( lcTemplate )

	*** Run with ExecScript
	CASE lnMode = 3
		RETURN loScript.ExecScript( lcTemplate, llIsString )    
ENDCASE

RETURN ""
ENDFUNC
*  wwUtils ::  RenderAspScript


************************************************************************
FUNCTION IsWinnt
*****************
***      Pass: llReturnVersionNumber
***    Return: .t. or .f.   or Version Number or -1 if not NT
*************************************************************************
LPARAMETER llReturnVersionNumber

loAPI=CREATE("wwAPI")
lcVersion = loAPI.ReadRegistryString(HKEY_LOCAL_MACHINE,;
           "SOFTWARE\Microsoft\Windows NT\CurrentVersion",;
           "CurrentVersion")
                          
IF !llReturnVersionNumber
  IF ISNULL(lcVersion)
     RETURN .F.
  ELSE
     RETURN .T.
  ENDIF
ENDIF

IF ISNULL(lcVersion)
   RETURN -1
ENDIF

RETURN VAL(lcVersion)   
* IsWinNt

************************************************************************
* wwUtils :: IsDotNet
****************************************
***  Function: Returns whether .Net is installed
***            Optionally returns the framework path and version
***            of the highest installed version.
***
***            Both parameters should be passed by ref
************************************************************************
FUNCTION IsDotNet(lcFrameworkPath,lcVersion)
LOCAL loAPI as wwAPI, x 

lcFrameworkPath = ""

loAPI = CREATEOBJECT("wwAPI")
lcWinDir = loAPI.getSystemdir(.T.)


*** Assume .NET 2.0
lcVersion = "v2.0.50727"
   
   *** Try to find the largest version number
lcFrameworkPath = lcWinDir + "Microsoft.NET\Framework\"
lnCount = ADIR(laNetDirs,lcFrameworkPath + "v?.*.*","D")
IF lnCount < 1
  lcVersion = ""
  lcFrameworkPath = ""
  RETURN .F.
ENDIF

*** Highest version comes last so go backwards through list
FOR x = lnCount TO 1 STEP -1
   lcVersion = laNetDirs[x,1]
   lcTPath = ADDBS(lcFrameworkPath + lcVersion )
   IF FILE(lcTPath + "regasm.exe")
     lcFrameworkPath = ADDBS(lcTPath)         
     EXIT
   ENDIF
ENDFOR   

RETURN .T.
ENDFUNC
*  wwUtils :: IsDotNet

**************************************************
FUNCTION IsCOMObject
*********************
*** Function: Checks to see if a COM object 
***           or ActiveX control exists
***   Assume: Uses wwAPI
***     Pass: lcProgId   - Prog Id of the Class
***           lcClassId  - (Optional) If passed in 
***                        by reference gets ClassId
***           lcClassDescript - (Optional) by ref
***   Return: .T. or .F.
*****************************************************
LPARAMETER lcProgId,lcClassId, lcClassDescript,lcServerPath

IF EMPTY(lcProgId)
   RETURN .F.
ENDIF

loAPI = CREATE("wwAPI")

*** Retrieve ClassId and Server Name
lcClassId =  ;
   loAPI.ReadRegistryString(HKEY_CLASSES_ROOT,;
                            lcProgId + "\CLSID",;
                            "")
IF ISNULL(lcClassId)
   lcClassId = ""
   lcClassDescription = ""
   RETURN .F.
ENDIF                                       

IF VARTYPE(lcClassDescript) = "C"                            
   lcClassDescript = ;
      loAPI.ReadRegistryString(HKEY_CLASSES_ROOT,;
                               lcProgId,"")

   IF ISNULL(lcClassDescript)
      lcClassDescript = ""
   ENDIF
ENDIF

IF VARTYPE(lcServerPath) = "C"
      lcServerPath = loAPI.ReadRegistryString(HKEY_CLASSES_ROOT,;
                               "CLSID\" + lcClassId + "\LocalServer32","")
      IF ISNULL(lcServerPath)
        lcServerPath = ""
      ENDIF                               
ENDIF       
                             
RETURN .T.

************************************************************************
FUNCTION RegisterOleServer
**************************
***  Function: Registers an OLE server or OCX control
***      Pass: lcServerPath  -  Full path and filename of OCX/OLE Server
***            llUnregister  -  .T. to unregister
***            llSilent      -  .T. or .F.
***    Return: .T. or .F.
*************************************************************************
LPARAMETER lcServerPath, llUnRegister, llSilent
LOCAL llRetVal, lcPath, lcOldPath

IF !FILE(lcServerPath)
    RETURN .f.
ENDIF

llRetVal=.F.
IF !llUnregister
   lcOldPath = SYS(5) + curdir()
   lcPath = JUSTPATH(lcServerPath)
   CD (lcPath)
   
   DECLARE INTEGER DllRegisterServer ;
      IN (lcServerPath)

   IF DllRegisterServer() = 0
      If !llSilent
         wait window nowait lcServerPath + " has been registered..."
      endif
      llRetVal=.T.
   ELSE
*!*         DECLARE INTEGER GETLASTERROR IN WIN32API
*!*         lnError = GetLastError()
*!*         wait window STR(lnError)
      wait window lcserverPath +  " could not be registered..." TIMEOUT 5
   ENDIF
   
   cd (lcOldPath)   
ELSE
   DECLARE INTEGER DllUnregisterServer ;
      IN (lcServerPath)

   IF DllUnregisterServer() = 0
      if !llSilent
        wait window nowait lcServerPath + " has been unregistered..."
      ENDIF
      llRetVal=.T.
   ENDIF
ENDIF      

RETURN llRetVal

#IF wwVFPVersion > 7
************************************************************************
* RegisterDotNetComponent
****************************************
***  Function: Registers a .Net COM component using RegAsm
***            by retrieving the runtime directory and executing
***            RegAsm from there.
***      Pass: lcDotNetDll - The .Net assembly to register
***            lcProgId    - One of the ProgIds to register
***                          used to check if the component registered
***                          "UNREGISTER" to unregister a component
***            @lcError    - Pass by reference in order to get error info
************************************************************************
FUNCTION  RegisterDotNetComponent(lcDotNetDLL as String, lcProgId as String,;
                                  lcError as String)  ;
                                  as Boolean
LOCAL lcFrameworkPath, lcVersion

IF lcProgId = "UNREGISTER"
   llUnregister = .T.
   lcProgId = ""
ELSE
   llUnregister = .F.
ENDIF

*** if the object is already registered exit
IF !EMPTY(lcProgId) AND ISCOMOBJECT(lcProgId)
   lcError = ""
   RETURN .T.
ENDIF

*** Try to register
lcFrameworkPath = ""
lcVersion = ""

IF !IsDotNet(@lcFrameworkPath)
   lcError = "DotNet Framework not installed or path not found."
   RETURN .F.
ENDIF

lcRun = ShortPath(ADDBS(lcFrameworkPath) + "regasm.exe") 
IF EMPTY(lcRun)  && File doesn't exist
   lcError = "Couldn't find RegAsm.exe at:" + CHR(13) +;
           lcFrameworkPath + "regasm.exe"
   RETURN .F.
ENDIF

IF llUnregister
   lcRun = lcRun + [ "] +  ShortPath(FULLPATH(lcDotNetDll)) + [" /unregister]
ELSE 
   lcRun = lcRun + [ "] +  ShortPath(FULLPATH(lcDotNetDll)) + [" /codebase]
ENDIF

WAIT WINDOW "Hang on. Trying to register " + ;
            JUSTFNAME(lcDotNetDLL) + " ..." + CHR(13) +;
            "This may take a few seconds..." NOWAIT

TRY
   *** Long command lines must run 'windows console'
   IF LEN(lcRun) > 115
      RUN /n &lcRun
      *** Wait long enough to start runtime and register
      *** this can be really slow so allow enough time
      INKEY(20)   
   ELSE
      RUN &lcRun
   ENDIF      
CATCH TO oExcept
    ASSERT .F. MESSAGE oExcept.Message
ENDTRY

WAIT CLEAR

IF EMPTY(lcProgId)
   RETURN .T.
ENDIF

llResult =  IsComObject(lcProgId)
IF !llResult
   lcError = "Registration of " + JUSTFNAME(lcDotNetDll) + " failed." + CHR(13) + CHR(13)+ ;
            "Command Line:" + CHR(13) + ;
            "RUN /n " + lcRun + CHR(13) + CHR(13) +;
            "Full deduced RegAsm Path:" + CHR(13) + ;
            lcFrameworkPath + "regasm.exe" + CHR(13) + CHR(13) +;
            "You can manually register your assembly by running REGASM.EXE" + CHR(13) + ;
            "from the framework BIN directory with the following command line: " + CHR(13)+;
            "<.Net framework bin path>\RegAsm /codebase <yourassembly>" + CHR(13) + CHR(13) + ;
            "The command line to register the component has been pasted into your ClipBoard"
            
ENDIF

RETURN llResult
* RegisterDotNetComponent
#ENDIF

************************************************************************
FUNCTION LaunchVsWebServer
****************************************
***  Function: Launches the Visual Studio Web Server
***    Assume:
***      Pass:  lnType = 0 - VS Web Server, 1 - Web Connection Cassini
***    Return:
************************************************************************
LPARAMETERS lcPhysical,lnPort, lcVirtual, lnType

IF VARTYPE(lnType) # "N"
  lnType = 0
ENDIF
  
IF EMPTY(lcPhysical)
   lcPhysical = ""
ENDIF
IF EMPTY(lnPort)
  lnPort = 81
ENDIF
IF EMPTY(lcVirtual)
  lcVirtual = "/"
ENDIF  

lcPath = ""
IF !IsDotNet(@lcPath)
	  RETURN .F.
ENDIF

IF lnType = 0
	lcPath = ADDBS(lcPath) + "webDev.WebServer.exe"
	IF RIGHT(lcPhysical,1) = "\"
		lcPhysical = SUBSTR(lcPhysical,1,LEN(lcPhysical)-1)
	ENDIF
	RETURN  CreateProcessEx(lcPath,[/port:] + TRANSFORM(lnPort) + [ /path:"] + lcPhysical + [" /vpath:"] + lcVirtual + ["],JUSTPATH(lcPath) )
ENDIF

lcPath = JUSTPATH(FULLPATH("console.exe")) + "\visualStudio\InternalWebServer\WebConnectionWebServer.exe"
RETURN  CreateProcessEx(lcPath,["] + lcPhysical + [" "] + TRANSFORM(lnPort) + [" "] +  lcVirtual + ["],JUSTPATH(lcPath) )	
ENDFUNC
*   LaunchVsWebServer

************************************************************************
FUNCTION DCOMCnfgServer
***********************
***  Function: Sets the security attributes of an Automation server
***            to Interactive User or a specific user account
***    Assume: If a password is passed DCOMPermissions.exe must be
***            available in the Foxpro path
***      Pass: lcProgId  -  Program ID for the server (wcdemo.wcdemoserver)
***            lcRunAs   -  User Account (Default: Interactive User)
***    Return: nothing
*************************************************************************
LPARAMETER lcProgId, lcRunAs, lcPassword
LOCAL lcProgId, loAPI, lcClassId, lcServerName

lcRunAs=IIF(type("lcRunAs")="C",lcRunAs,"Interactive User")
lcProgId=IIF(type("lcProgId")="C",lcProgId,"")
lcPassword=IIF(EMPTY(lcPassword),"",lcPassword)


loAPI = CREATE("wwAPI")

*** Retrieve ClassId and Server Name
lcClassId = loAPI.ReadRegistryString(HKEY_CLASSES_ROOT,;
                                     lcProgId + "\CLSID",;
                                     "")
lcServerName = loAPI.ReadRegistryString(HKEY_CLASSES_ROOT,;
                                        lcProgId + "","")

IF ISNULL(lcClassId) or ISNULL(lcServerName)
  wait window nowait "Invalid Class Id..."
  RETURN
ENDIF
  
wait window "Configuring server security for "+CR+;
             lcProgId + CR + lcServerName  nowait

*** Check whether CLSID/APPID exists
lcAppIdKey = loAPI.ReadRegistryString(HKEY_LOCAL_MACHINE,;
				            "AppId\"+lcClassId,"RunAs",.f.)

IF !EMPTY(lcPassword) 
   *** If lcAppId Key doesn't exist create the key 
   *** by configuring the server to run under SYSTEM account
   *** first (which manually writes out the key)
   *** Otherwise the called EXE will fail
   IF ISNULL(lcAppIdKey) OR EMPTY(lcAppIdKey)	
      DCOMCnfgServer(lcProgId,"SYSTEM")
   ENDIF

   IF !FILE("dcompermissions.exe")
      WAIT WINDOW "Couldn't find dcompermissions.exe..." TIMEOUT 5
      RETURN
   ENDIF
   
   lcPath = FULLPATH("dcompermissions.exe")
   lcPath = SHORTPATH(lcPath)

   lcCmd = ;
      "RUN " +lcPath +" -runas " + lcClassId + " " + lcRunas + " " + lcPassword +" > dcom.txt"
       
   &lcCMD

   lcResult = FILETOSTR("DCOM.TXT")
   ERASE DCOM.TXT

   IF !EMPTY(lcResult) AND ATC("ERROR:",lcResult) >0
      MESSAGEBOX(lcProgId + CHR(13) + "Account: " + lcRunas + CHR(13)+ CHR(13) +lcResult ,48,"DCOM Permissions")
      RETURN .F.
   ENDIF
ELSE
   *** Now add AppId key to the ClsID entry
   if !loAPI.WriteRegistryString(HKEY_LOCAL_MACHINE,;
           "SOFTWARE\Classes\CLSID\"+lcClassId,"AppId",lcClassID,.t.)
      wait window "Unable to write AppID value..."  nowait
      RETURN
   ENDIF                                     

   *** Create a AppID Entry if it doesn't exist
   if !loAPI.WriteRegistryString(HKEY_CLASSES_ROOT,;
           "AppID\"+lcClassId,CHR(0),CHR(0),.t.)
      wait window "Unable to write AppID key..."  nowait
      RETURN
   ENDIF                                     

   *** Write the Server Name into the Default key
   loAPI.WriteRegistryString(HKEY_CLASSES_ROOT,;
                             "AppID\"+lcClassId,"",;
                             lcServerName,;
                             .t.)

   *** Write Interactive User (or user Accounts)                          
   loAPI.WriteRegistryString(HKEY_CLASSES_ROOT,;
                             "AppID\"+lcClassId,"RunAs",;
                             lcRunAS,;
                             .t.)
ENDIF

wait window "DCOM security context set to: " + lcRunAs nowait                          
RETURN 


************************************************************************
* wwUtils :: DCOMLaunchPermissions
****************************************
***  Function: Sets DCOM Launch and Access Permissions for a specific
***            server. If no ProgId is passed in Default permissions are 
***            set.
***    Assume: DCOMPermissions.exe is in Fox path
***      Pass: lcProgId   -  Progid of COM object
***            lcUserName -  Username to add to Access/Launch permissions
***            @lcErrorMsg - Pass by reference to get error message back
***    Return: .T.  or .F.     (UI shows error messages)
************************************************************************
FUNCTION DCOMLaunchPermissions
LPARAMETERS lcProgid,lcusername, lcErrorMsg

IF !EMPTY(lcProgid)
   lcClassID=""
   llResult = IsComObject(lcProgId,@lcClassID)

   IF EMPTY(lcClassID)
      WAIT WINDOW "Invalid Prog ID"
      RETURN .F.
   ENDIF
ELSE
   lcClassId = ""
ENDIF

lcPath = FULLPATH("dcompermissions.exe")
lcPath = SHORTPATH(lcPath)


IF EMPTY(lcClassID)
   *** Set Default permissions
   lcCmd = ;
   "RUN " +lcPath +" -da " + lcClassId + [ set "] + lcUserName + [" permit > dcom.txt]
   _cliptext = lcCMD
ELSE
   lcCmd = ;
   "RUN " +lcPath +" -aa " + lcClassId + [ set "] + lcUserName + [" permit > dcom.txt]
   _cliptext = lcCMD
ENDIF
&lcCMD

lcResult = FILETOSTR("DCOM.TXT")
ERASE DCOM.TXT

IF !EMPTY(lcResult) AND ATC("ERROR:",lcResult) > 0
   MESSAGEBOX(lcProgId + CHR(13) + "Account: " + lcUserName + CHR(13)+ CHR(13) +lcResult ,48,"DCOM Permissions")
   RETURN .F.
ENDIF

IF EMPTY(lcClassID)
   lcCmd = ;
   "RUN " +lcPath +" -dl " + lcClassId + [ set "]  + lcUserName + [" permit > dcom.txt]
   _cliptext = lcCMD
ELSE
   lcCmd = ;
   "RUN " +lcPath +" -al " + lcClassId + [ set "] + lcUserName + [" permit > dcom.txt]
   _cliptext = lcCMD   
ENDIF

&lcCMD

lcResult = FILETOSTR("DCOM.TXT")
ERASE DCOM.TXT

IF !EMPTY(lcResult) AND ATC("ERROR:",lcResult) >0
   MESSAGEBOX(lcProgId + CHR(13) + "Account: " + lcUserName + CHR(13)+ CHR(13) +lcResult ,48,"DCOM Permissions")
   RETURN .F.
ENDIF

RETURN
ENDFUNC

************************************************************************
* wwUtils :: SetACL
****************************************
***  Function:
***    Assume:
***      Pass: lcPath        -   Path or Filename to assign ACE to
***            lcUser        -   The Username to assign
***            lcAccess      -   [N]one, [R]ead, [C]hange, [F]ull
***            llInherit     -   Pass rights down the directory tree
***            llReplace     -   Replace rights on the resource
***                              (deletes all ACL entries except this one)
***    Return: 
************************************************************************
FUNCTION SetACL(lcPath,lcUser,lcAccess,llInherit,llReplace)
LOCAL lcCommand, lcFile, lcParms

IF EMPTY(lcUser)
   RETURN .T.
ENDIF

*** Strip off any trailing backslashes
IF RIGHT(lcPath,1)="\" AND LEN(lcPath) > 1
   lcPath = SUBSTR(lcPath,1,LEN(lcPath)-1)
ENDIF

lcParms = ShortPath( lcPath) 
IF llInherit
   lcParms = lcParms + " /T "
ENDIF
IF !llReplace 
   lcParms = lcParms + " /E " 
ENDIF
lcParms = lcParms + " /P " + lcUser + ":" + lcAccess


*** Use CreateProcessEx if wwIPStuff is available
*** so we don't get flashing DOS Windows
IF FILE("wwipstuff.dll")
   LOCAL oAPI
   oAPI = CREATEOBJECT("wwAPI")
   CreateProcessEx( oAPI.GetSystemDir() + "cacls.exe",ALLTRIM(lcParms),,2,.t.,FULLPATH("cacls.txt"))
ELSE
   lcCommand = "RUN Cacls.exe " 
   lcCommand = lcCommand + " " + lcParms + " > cacls.txt"
   &lcCommand
ENDIF   

lcFile = ""
IF FILE("Cacls.txt")
   lcFile = FILETOSTR("cacls.txt")
   ERASE Cacls.txt
ENDIF
IF EMPTY(lcFile)
   RETURN .F.
ENDIF

RETURN .T.
ENDFUNC
*  wwUtils :: SetACL


************************************************************************
FUNCTION CopyTree
*****************
***  Function: Copies a directory tree with wild card expression.
***            Any files in subdirectories are copied and all 
***            directories including empty ones are created.
***      Pass: lcSource -  Source Path and File Template (c:\temp\*.*)
***            lcTarget -  Target Directory to move trees to
***    Return: nothing
*************************************************************************
LPARAMETERS lcSource, lcTarget, llNoRecursion
LOCAL x, lcSourcePath,lcTargetPath, lcSourceFiles

lcSourcePath = ADDBS(JUSTPATH(lcSource))
lcTargetPath = ADDBS(lcTarget)
lcSourceFiles = JUSTFNAME(lcSource)
*lcTargetFiles = JUSTFNAME(lcTarget)

lcAlias = ALIAS()
SELECT 0
CREATE CURSOR TDirs( Path c(128) )

IF !llNoRecursion
	GetDirs(ADDBS(lcSourcePath),lcSourcePath)
ELSE
    *** Insert empty relative path
    INSERT INTO TDirs  (Path) VALUES ("")
ENDIF	

SCAN
   IF !Directory(lcTargetPath + TRIM(path))
      MD (lcTargetPath + TRIM(Path))
   ENDIF
    
   *** Check for files first
   IF EMPTY(SYS(2000,lcSourcePath + TRIM(Path) + lcSourceFiles) )
      LOOP
   ENDIF
    
   COPY FILES (lcSourcePath + TRIM(Path)  + lcSourceFiles) TO ;
              (lcTargetPath + TRIM(Path) + lcSourceFiles)
ENDSCAN

USE 

IF !EMPTY(lcAlias) AND USED(lcAlias)
   SELECT (lcALias)
ENDIF

RETURN

************************************************************************
* DeleteTree
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION DeleteTree(lcSource)
PRIVATE lcSourcePath
LOCAL llError, loEval, lcCode, lcAlias

lcSourcePath = ADDBS(lcSource)
llError = .F.

lcCurPath = SYS(5)+CURDIR()
lcAlias = ALIAS()
CREATE CURSOR TDirs( Path c(128) )


GetDirs(ADDBS(lcSourcePath),lcSourcePath)
INDEX ON Path DESCENDING TAG PATH
IF _TALLY < 1
   RETURN .F.
ENDIF

SCAN
   lcPath = lcSourcePath + TRIM(path)
   ERASE (lcPath + "*.*")
   IF ISDir(lcPath)
      RD (lcPath)
   ENDIF
ENDSCAN


RETURN !llError
ENDFUNC
* DeleteTree

*******************************
FUNCTION GetDirs(lcPath,lcBase)
*******************************
LOCAL x, lnFiles, laDirs[1]

*** Force path to be proper cased
lcPath = LOWER(lcPath)
lcBase = LOWER(lcBase)

INSERT INTO TDirs (Path) Values ( STRTRAN(lcPath,lcbase,"",-1,-1,1) )

*** Get directories first
lnFiles = aDir(laDirs,ADDBS(lcPath) + "*.*" ,"D",1)
FOR x=1 to lnFiles
   if  AT("D",laDirs[x,5]) = 0 OR laDirs[x,1] = "." OR laDirs[x,1] = "    "
      LOOP
   ENDIF
   GetDirs( ADDBS(lcPath + laDirs[x,1]),lcBase ) 
ENDFOR
ENDFUNC





************************************************************************
* wwUtils :: SpellCheck
****************************************
***  Function: Uses Word's Spell Check functionality to spell check
***            a string of text interactively
***    Assume: Requires Word 2000 and IE 4 or later COM objects
***      Pass:  lcText    - the text to spell check
***             llIsHtml  - if .t. the text is HTML
***    Return: Spellchecked text
************************************************************************
FUNCTION SpellCheck(lcText,llIsHtml,llHideForm)
LOCAL loWord, loDoc, x, y, llActiveForm

IF EMPTY(lcText)
   RETURN ""
ENDIF

IF !ISCOMOBJECT("Word.Application")
   RETURN lcText
ENDIF   

IF !llIsHtml
   *** Plain Text - simple assign and retrieve
   LOCAL loWord as Word.Application
   loWord = CREATEOBJECT("Word.Application")

   LOCAL loDoc as Word.Document
   loDoc = loWord.Documents.Add(,,1,.T.)	
   loDoc.Content.Text = lcText

   loWord.Visible = .T.

   llActiveForm = TYPE("_VFP.ActiveForm") = "O"    
   LOCAL loForm 
   loForm = null
   IF llActiveForm
      loForm = _VFP.ActiveForm
   ENDIF
     
   IF llActiveForm AND llHideForm
     _VFP.ActiveForm.WindowState = 1
   ENDIF
      
   loDoc.CheckSpelling()

   IF llActiveForm AND llHideForm AND !ISNULL(loForm)
	   loForm.WindowState = 0
   ENDIF
      
   lcText = loDoc.Content.Text
   
   loWord.Visible = .f.
   loDoc.Close(.f.)
   loDoc = .null.

   loWord.Quit(.f.)
   loWord = .null.   

   RETURN lcText
ENDIF

*** HTML - load into IE, retrieve text, replace 
***        changed text
loIE = CREATEOBJECT("InternetExplorer.Application")
loIE.Navigate("about:blank")

DO WHILE loIE.Busy
   DOEVENTS
ENDDO

loIEDoc = loIE.Document
loIEDoc.Body.innerHtml = lcText

lcTText = loIEDoc.Body.innerText
loIE = .f.

loWord = CREATEOBJECT("Word.Application")
*loWord.WindowState= 2  && wdWindowStateMinimize

loDoc = loWord.Documents.Add(,,1,.T.)

loDoc.Content.Text = lcTText

*** Pick up all the error text
x = 0
FOR EACH loError IN loDoc.SpellingErrors
     x = x + 1
     DIMENSION laErrors[X,2]
     laErrors[x,1] = loError
     laErrors[x,2] = loError.Text && Old text
ENDFOR

IF x > 0
   loDoc.SpellingErrors.Item(1).CheckSpelling()
ENDIF
 
FOR y = 1 TO  x
    IF laErrors[y,1].Text # laErrors[y,2]
       lcText = STRTRAN(lcText,laErrors[y,2],laErrors[y,1].Text)
    ENDIF
    laErrors[y,1] = .null.
ENDFOR

loWord.Visible = .f.
loDoc.Close(.f.)
loDoc= .null.
loWord.Quit(.f.)
loWord = .null.

RETURN lcText
*  wwUtils :: SpellCheck



#IF .F.
*ENDDEFINE
#ENDIF

*** Commonly used UI Functions ***

* Pass lnX/YFactor by reference
FUNCTION TwipsFactor(lnXFactor, lnYFactor)
LOCAL ln_x_pixels, ln_y_pixels, ln_twips, ln_partial_x, ln_partial_y, ;
     ln_hwnd, ln_hdc

*** Calculate the factor to be used in the HitTest method...
ln_x_pixels = 88
ln_y_pixels = 90
ln_twips    = 1440

DECLARE INTEGER GetActiveWindow IN win32api
DECLARE INTEGER GetActiveWindow IN win32api
DECLARE INTEGER GetDC           IN win32api INTEGER iHDC
DECLARE INTEGER GetDeviceCaps   IN win32api INTEGER iHDC, INTEGER iIndex

ln_hwnd = GetActiveWindow()
ln_hdc  = GetDC(ln_hwnd)

ln_partial_x = GetDeviceCaps(ln_hdc, ln_x_pixels)
ln_partial_y = GetDeviceCaps(ln_hdc, ln_x_pixels)

lnXFactor = ln_twips/ln_partial_x
lnYFactor = ln_twips/ln_partial_y
RETURN

************************************************************************
* wwUtils ::  IsAdmin
****************************************
***  Function: Determines whether user is an admin user by probing
***            access to HKLM registry key
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION IsAdmin()
LOCAL loAPI, lcVal

DECLARE INTEGER IsUserAnAdmin IN Shell32
TRY
	lnResult = IsUserAnAdmin()
CATCH
	*** OLD OLD Version of Windows assume .T.
	lnResult = 1
ENDTRY
IF lnResult = 0
   RETURN .F.
ENDIF

RETURN .T.  
ENDFUNC
*  wwUtils ::  IsAdmin

************************************************************************
* wwUtils :: GetSystemPassword
****************************************
***  Function: Retrieves the AutoLogon Password
***    Assume: Used for demos so I don't have to type my pass
***      Pass: optional llUsername - returns username
***    Return: string
************************************************************************
FUNCTION GetSystemPassword
LPARAMETERS llUserName

loAPI = NEWOBJECT("wwAPI","wwapi.prg")
IF !llUserName
   lcPass=loAPI.ReadRegistryString("HKLM","SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","DefaultPassword")
ELSE
   lcPass=loAPI.ReadRegistryString("HKLM","SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","DefaultUserName")
ENDIF

IF ISNULL(lcPass)
   RETURN ""
ENDIF

RETURN lcPass
ENDFUNC
*  wwUtils :: GetSystemPassword

************************************************************************
* wwUtils :: PasswordForm
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetPassword(lcMessage)
IF EMPTY(lcMessage)
  lcMessage = "Please enter your password"
ENDIF
RETURN InputForm(SPACE(20),lcMessage,"Password Entry",,,"PASSWORD")
ENDFUNC
*  wwUtils :: PasswordForm


************************************************************************
FUNCTION InputForm
******************
***  Function: Creates a simple Input form that returns a value
***    Assume: Consists of this function and Form Class
***      Pass: lcValue   -   Initial value to retrieve
***            lcMessage -   The request message
***            lcCaption -   Form Caption (_Screen.Caption)
***            lnFormWidth   Width of the form
***            lnFieldWidth  Widht of the input field
***            lcFormat		  Format string for the input field
***                          "PASSWORD" to show password entry
***            lcCancelValue Value returned on a Cancel operation 
***    Return: Value or ("" or -1)
************************************************************************
LPARAMETER lcValue, lcMessage, lcCaption, lnFormWidth, lnFieldWidth, lcFormat,lcCancelValue
PRIVATE pcResult
LOCAL o

IF PCOUNT() > 6
  pcCancelValue = lcCancelValue
ELSE
  pcCancelValue = NULL
ENDIF

lcValue=IIF(EMPTY(lcValue),"",lcValue)
lcMessage=IIF(EMPTY(lcMessage),"Please enter",lcMessage)
lcCaption=IIF(EMPTY(lcCaption),_SCREEN.Caption,lcCaption)
lnFormWidth=IIF(EMPTY(lnFormWidth),300,lnFormWidth)
lnFieldWidth=IIF(EMPTY(lnFieldWidth),lnFormWidth - 20,lnFieldWidth)
lcFormat=IIF(EMPTY(lcFormat),"@K",lcFormat)


lcType = VARTYPE(lcCancelValue)
pcResult = lcValue

o=CREATE("frmInput")
o.Width = lnFormWidth
o.nFieldWidth = lnFieldWidth
o.Caption = lcCaption
o.cMessage = lcMessage
o.cFormat = lcFormat

IF lcFormat = "PASSWORD"
   o.cFormat = "@K"
   o.txtInput.Fontname="Symbol"
   o.txtInput.PasswordChar = "*"
ENDIF

o.Show()

IF TYPE("pcResult")="C"
   lcValue = TRIM(pcResult)
ELSE
   lcValue = pcResult
ENDIF   
RETURN lcValue


**************************************************
*-- Form:         frminput 
*-- ParentClass:  form
*-- BaseClass:    form
DEFINE CLASS frminput AS form
    nFieldWidth = 250
    cMessage = "Please enter:"
    cFormat = ""

	Top = 0
	Left = 0
	Height = 90
	Width = 300
	ControlBox = .F.
	Name = "frmInput"
	WindowType = 1
	AutoCenter = .t.
	Showwindow = 1
	BorderStyle = 2
	MinButton = .f.
	ShowWindow = 1
	MaxButton = .f.

	ADD OBJECT lblMessage AS label WITH ;
		AutoSize = .T., ;
		Caption = "Message Text:", ;
		Height = 17, ;
		Left = 7, ;
		Top = 11, ;
		Width = 81, ;
		Name = "lblMessage",;
		Font = "Tahoma" ,;
		FontSize = 8
		
	ADD OBJECT txtinput AS textbox WITH ;
		ControlSource = "pcResult", ;
		Height = 22, ;
		Left = 5, ;
		Top = 28, ;
		Width = 373, ;
		Name = "txtInput",;
		Font = "Tahoma" ,;
        Default = .T.,;
		FontSize = 8
	ADD OBJECT cmdOk AS commandbutton WITH ;
		Top = 55, ;
		Left = THISFORM.width - 150, ;
		Height = 25, ;
		Width = 70, ;
		Caption = "OK", ;
		Default = .T., ;
		Fontname="Tahoma",;
		Fontsize = 8,;
        Cancel = .F.,;
		Name = "cmdOK"

	ADD OBJECT cmdCancel AS commandbutton WITH ;
		Top = 55, ;
		Left = THISFORM.Width - 75, ;
		Height = 25, ;
		Width = 70, ;
		Caption = "\<Cancel", ;
		Fontname="Tahoma",;
		Fontsize = 8,;
		Cancel = .T.,;
		Name = "cmdCancel" 

    PROCEDURE Show
    	this.AlwaysOnTop = .t.    	 
        THIS.Icon = _Screen.icon
        THIS.txtInput.Width = THIS.nFieldWidth
        this.txtInput.MaxLength = 250
        THIS.lblMessage.Caption = THIS.cMessage
		THIS.cmdCancel.Left = THISFORM.Width - 85
		THIS.cmdOk.Left = THISFORM.Width - 157
        IF !EMPTY(THIS.cFormat)
           IF ATC("@",THIS.cFormat) > 0
             THIS.txtInput.Format = THIS.cFormat
           ELSE
             THIS.txtInput.InputMask = THIS.cFormat
           ENDIF
        ENDIF
        DoEvents
        this.AlwaysOnTop = .F.         
    ENDPROC
	PROCEDURE cmdOk.Click
		RELEASE THISFORM
	ENDPROC
	PROCEDURE cmdCancel.Click
	    LOCAL lcType
       
       IF ISNULL(pcCancelValue)
   	    lcType = TYPE("pcResult")
          DO CASE
               CASE lcType $ "CM"
       	 	    pcResult = ""
        	    CASE lcType $ "NIBY"
            	    pcResult = -99999999
            	CASE lcType $ "DT"
            	    pcResult = {}
            	CASE lcType = "L"
            	    pcResult = .f.
           ENDCASE
      ELSE
         pcResult = pcCancelValue
      ENDIF
	   RELEASE THISFORM
	ENDPROC
ENDDEFINE
*
*-- EndDefine: frminput
**************************************************

#if wwVFPVersion > 6
*************************************************
*PROCEDURE Qk_Menu
******************
***   Author: (c) Rick Strahl, 1994
*** Modified: 1/10/05
*** Function: Creates 'quick' popup menu at the 
***           current mouse pointer location. 
***           Can optionally be centered or 
***           placed at custom screen location.
***
***   Assume: parameters passed as pairs of
***           Prompt/Command. requires VFP7 or later
***
***     Pass: 
***     pcArg1 - 32 Menu bar text
***     pcArg1 - Comma separated list of bar text strings
*** 
***  Return: "" if nothing is clicked
***          otherwise the prompt is returned
***          \< is stripped from the prompt!
*************************************************
FUNCTION QK_Menu
PARAMETERS pcArg1,pcArg2,pcArg3,pcArg4,pcArg5,;
   pcArg6,pcArg7,pcArg8,pcArg9,pcArg10,;
   pcArg11,pcArg12,pcArg13,pcArg14,pcArg15,;
   pcArg16,pcArg17,pcArg18,pcArg19,pcArg20,;
   pcArg21,pcArg22,pcArg23,pcArg24,pcArg25,;
   pcArg26,pcArg27,pcArg28,pcArg29,pcArg30,;
   pcArg31,pcArg32

PRIVATE lnParms,x,a_parms,lcPrompt,lcCommand,;
   lcPromptVal,lcCommandVal,lcOldRMouse,lnBarno, lncount

*** Assume all options are pairs
lnParms=PARAMETERS()
IF lnParms < 1
   RETURN ""
ENDIF

*** Default to Mouse Position
IF TYPE("pnLeft") = "N"
   lnCol = pnLeft 
   lnRow = pnTop
ELSE
   lnRow=MROW()   && Location for popup in Screen
   lnCol=MCOL()
ENDIF

lnMaxLen=0      && Longest longest prompt string

lcWindow = WONTOP()

IF !EMPTY(lcWindow)
   *** Now activate the popup
   DEFINE POPUP qk_menu IN WINDOW (WONTOP()) ;
      FROM lnRow,lnCol ;
      MARGIN SHORTCUT
ELSE
   *** Now activate the popup
   DEFINE POPUP qk_menu  ;
      FROM lnRow,lnCol ;
      MARGIN SHORTCUT
ENDIF

DIMENSION a_parms[lnParms]

IF lnParms = 1 AND AT(",",pcArg1) > 0
   lnCount = ALINES(a_Parms,pcArg1,0,",")
   FOR x = 1 TO lnCount 
      DEFINE BAR x OF qk_menu  PROMPT a_Parms[x]
   ENDFOR
ELSE
   *** Build array of the arguments 
   *** so it's easier to deal with
   FOR x=1 TO lnParms
      *** Build the variable name
      lcPrompt="pcArg"+TRANSFORM(x)

      *** Save result values 
      *** to minmize & and EVALUATEs
      lcPromptVal=EVALUATE(lcPrompt)
      
      a_parms[x] = lcPromptVal

      DEFINE BAR x OF qk_menu  PROMPT lcPromptVal
   ENDFOR &&* x=1 to lnParms
ENDIF

ON SELECTION POPUP QK_MENU DEACTIVATE POPUP QK_MENU
   
ACTIVATE POPUP QK_MENU
RELEASE POPUP QK_MENU

lnBarno=bar()

IF lnBarNo = 0
   RETURN ""
ENDIF

RETURN STRTRAN(a_parms[lnBarNo],"\<","")
*EOP Qk_Menu
#ENDIF




