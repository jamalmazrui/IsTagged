' IsTagged
' Version 1.0
' November 5, 2015
' Copyright 2012 - 2015 by Jamal Mazrui
' GNU Lesser General Public License (LGPL)

#INCLUDE "C:\IsTagged\Private.inc"
#INCLUDE "Win32API.inc"

#Resource RCDATA, 1, "IsTagged_temp.dll"
#Resource RCDATA, 2, "IsTagged_temp.tlb"

FUNCTION FileToString(BYVAL s_file AS ASCIIZ * 256) AS STRING
LOCAL i_size AS LONG, h_file AS LONG, s_return AS STRING

IF LEN(DIR$(s_file, 7)) =0 THEN
s_return =""
ELSE
h_file =FREEFILE
OPEN s_file FOR BINARY AS h_file
i_size =LOF(h_file)
GET$ h_file, i_size, s_return
CLOSE h_file
END IF
FUNCTION =s_return
END FUNCTION

FUNCTION StringToFile(BYVAL s_text AS STRING, BYVAL s_file AS ASCIIZ * 256) AS LONG
LOCAL i_size AS LONG, h_file AS LONG, i_return AS LONG

IF ISTRUE ISFILE(s_file) THEN KILL s_File
'msgbox format$(len(s_text))
IF LEN(s_text) =0 THEN
'If IsFalse Then
i_return =0
ELSE
h_file =FREEFILE
OPEN s_file FOR BINARY AS h_file
i_size =LEN(s_text)
PUT$ h_file, s_text
CLOSE h_file
i_return =1
END IF
FUNCTION =i_return
END FUNCTION

FUNCTION PrintLine(Z AS STRING) AS LONG
 ' returns TRUE (non-zero) on success
   LOCAL hStdOut AS LONG, nCharsWritten AS LONG
   LOCAL w AS STRING
   STATIC CSInitialized AS LONG, CS AS CRITICAL_SECTION
   IF ISFALSE CSInitialized THEN
       InitializeCriticalSection CS
       CSInitialized  =  1
   END IF
   EntercriticalSection Cs
   hStdOut      = GetStdHandle (%STD_OUTPUT_HANDLE)
   IF hSTdOut   = -1& OR hStdOut = 0&   THEN     ' invalid handle value, coded in line to avoid
                                                 ' casting differences in Win32API.INC
                                                 ' test for %NULL added 8.26.04 for Win/XP
     AllocConsole
     hStdOut  = GetStdHandle (%STD_OUTPUT_HANDLE)
   END IF
   LeaveCriticalSection CS
   w = Z & $CRLF
   FUNCTION = WriteFile(hStdOut, BYVAL STRPTR(W), LEN(W),  nCharsWritten, BYVAL %NULL)
 END FUNCTION

FUNCTION StringPlural(sText AS STRING, iCount AS LONG) AS STRING
LOCAL sReturn AS STRING

sReturn = sText
IF iCount <> 1 THEN sReturn = sReturn & "s"
FUNCTION = sReturn
END FUNCTION

FUNCTION DialogInput(sTitle AS STRING, sMessage AS STRING, sValue AS STRING) AS STRING
FUNCTION = INPUTBOX$(sMessage, sTitle, sValue)
END FUNCTION

FUNCTION DialogShow(sTitle AS STRING, sMessage AS STRING) AS LONG
' show a standard message box

DIM iFlags AS LONG

DialogShow = %True
iFlags = %MB_ICONINFORMATION OR %MB_SYSTEMMODAL
IF sTitle = "" THEN sTitle = "Show"
MSGBOX sMessage, iFlags, sTitle
END FUNCTION

FUNCTION StringQuote(BYVAL s$) AS STRING
FUNCTION = CHR$(34) & s$ & CHR$(34)
END FUNCTION

FUNCTION DialogConfirm(sTitle AS STRING, sMessage AS STRING, sDefault AS STRING) AS STRING
' Get choice from a standard Yes, No, or Cancel message box

DIM iFlags AS LONG, iChoice AS LONG

DialogConfirm = ""
iFlags = %MB_YESNOCANCEL
iFlags = iFlags OR %MB_ICONQUESTION     ' 32 query icon
iFlags = iFlags OR %MB_SYSTEMMODAL ' 4096   System modal
IF sTitle = "" THEN sTitle = "Confirm"
IF sDefault = "N" THEN iFlags = iFlags OR %MB_DEFBUTTON2
iChoice = MSGBOX(sMessage, iFlags, sTitle)
IF iChoice = %IDCANCEL THEN EXIT FUNCTION

IF iChoice = %IDYES THEN
DialogConfirm = "Y"
ELSE
DialogConfirm = "N"
END IF
END FUNCTION

Function GetTempDir() As String
Local zPath As ASCIIZ*%MAX_PATH
GetTempPath(%MAX_Path, zPath)
Function = zPath
End Function

Function PBMain() As Long
Local hDoc As DWord
Local iResult As Long
Local sBase, sBody, sDll, sExt, sClsId, sFile, sLib, sPassword, sPath, sPdf, sRoot, sSpec, sStatus, sTlb, sUnlockKey As String
Local oLib As Dispatch
Local vPassword, vPdf, vUnlockKey As Variant

sSpec = Command$
sSpec = Trim$(sSpec, ANY $Spc + $Dq)
IF Len(sSpec) = 0 Then
PrintLine("Syntax:" + $CrLf + "IsTagged.exe FileSpec")
EXIT FUNCTION
END IF

PrintLine("Checking Tagged status of " + sSpec)

sLib = "IsTagged_temp.dll"
sDll = GetTempDir() + "IsTagged_temp.dll"
sBody = Resource$(RCDATA, 1)
StringToFile(sBody, sDll)
sTlb = GetTempDir() + "IsTagged_temp.tlb"
sBody = Resource$(RCDATA, 2)
StringToFile(sBody, sTlb)

sLib = sDll
sClsId = GUID$("{2E75DB15-9312-4902-8DA0-EAC34A6DD40C}")
oLib = NewCom ClsId sClsId Lib sLib
sUnlockKey = $QuickPDF_UnlockKey
vUnlockKey = sUnlockKey
Object Call oLib.UnlockKey(vUnlockKey) To iResult

sPdf = Dir$(sSpec)
While Len(sPdf) > 0
' sPdf = PathScan$(FULL, sPdf)
sPdf = PathName$(PATH, sSpec) + sPdf
sPdf = PathScan$(FULL, sPdf)
vPdf = sPdf
Object Call oLib.LoadFromFile(vPdf, vPassword) To iResult

If IsFalse iResult Then
Object call oLib.LastPageError() To iResult
' DialogShow("LastPageError", Format$(iResult))
sStatus = "O"
Else
sStatus = "N"
Object Call oLib.SelectedDocument() To hDoc

Object Call oLib.IsTaggedPDF() To iResult
If IsTrue iResult Then sStatus = "Y"
Object Call oLib.RemoveDocument(hDoc) To iResult
End If ' LoadFromFile result

PrintLine(sStatus + " " + sPdf)
sPdf = Dir$()
WEnd

' Object Call oLib.CloseFile(hPdf) To iResult
Object Call oLib.ReleaseLibrary() To iResult
OLib = Nothing

If IsFile(sDll) Then Kill sDll
If IsFile(sTlb) Then Kill sTlb
End Function
