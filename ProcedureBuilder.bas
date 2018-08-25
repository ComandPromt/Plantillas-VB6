Attribute VB_Name = "mProcedureBuilder"
Option Explicit
Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
'====================================================================
'this sub should be executed from the Immediate window
'in order to get this app added to the VBADDIN.INI file
'you must change the name in the 2nd argument to reflecty
'the correct name of your project
'====================================================================
Sub AddToINI()
    Dim ErrCode As Long
    ' Add the add-in into VBADDINI.INI
    ErrCode = WritePrivateProfileString("Add-Ins32", "ProcedureBuilder.Connect", "0", _
    "vbaddin.ini")
    MsgBox "Add-in is now entered in VBADDIN.INI file."
    ' Write me to the system registry for the first run
    SaveSetting "Procedure Builder", "Author Details", "Author", "Mark Kirkland"
    SaveSetting "Procedure Builder", "Author Details", "Organisation", "Brighton Health Care NHS Trust"
End Sub
Function NoSpaces(SearchString As String) As Boolean

' Function searches a string for the space character and returns TRUE if no spaces are found

' Error Handler
On Error GoTo NoSpaces_Error

' Declare variables
Dim Counter As Integer  ' Loop counter
Dim strTestChar As String

For Counter = 1 To Len(SearchString)
    ' Search the string a character at a time
    strTestChar = Mid$(SearchString, Counter, 1)
    ' If a space character is found then return FALSE and exit the function
    If strTestChar = " " Then
        NoSpaces = False
        Exit Function
    End If
Next Counter

' No spaces were found so return TRUE
NoSpaces = True
Exit Function

' Error Routine
NoSpaces_Error:
MsgBox "Error: " & Err.Number & " - " & Err.Description & "NoSpaces"

End Function
