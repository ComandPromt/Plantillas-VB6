Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Const MAX_PATH& = 260


Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long


Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long


Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long


Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long


Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long


Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public presets(1 To 5) As String
Public recreate(1 To 5) As String
Type vz
    winamp_path As String
    mdload_delay As Integer
End Type
Public d As vz
Public sfile As String

Sub Main()
sfile = App.Path + "\wac.ini"
getSettings
'MsgBox isWinampRunning(d.winamp_path)

If d.winamp_path = "" Then
    frmOptions.Show
Else
    frmSplash.Show
End If
End Sub
Public Function returnED(truthVal As Boolean) As String
If truthVal = False Then
    returnED = "Disabled"
Else
    returnED = "Enabled"
End If
End Function

Public Function FormatTime(InSeconds As Integer) As String
    Dim m As Integer, s As Integer
    m = Int(InSeconds \ 60)
    s = InSeconds Mod 60
    FormatTime = Trim(Str(m) + ": " + Format(s, "00"))
End Function

Function Parse(ByVal parseStringx, ByVal argNum As Integer) As Variant
On Error Resume Next
Dim lastPos As Integer
Dim subPos As Integer
Dim argPos(1 To 50) As Integer
Dim argContent(1 To 50)
Dim parsestring As String
Dim argcount As Integer
parsestring = parseStringx
parsestring = Trim(Right(parsestring, ((Len(parsestring)) - (InStr(parsestring, " ")))))

parsestring = parsestring & " " 'save my ass some work



'count arguments
argcount = 0
Do
    DoEvents
    lastPos = InStr((lastPos + 1), parsestring, " ")
    If lastPos = 0 Then Exit Do
    argcount = argcount + 1
    argPos(argcount) = lastPos
Loop
If argcount = 0 Then Exit Function
'end count arguments

'get argument content
Dim i As Integer
For i = 1 To argcount
    Select Case i
        Case argcount
            If argcount <> 1 Then
                subPos = argPos(i - 1)
            Else
                subPos = 1
            End If
        Case 1
            subPos = 1
        Case Else
            subPos = argPos(i - 1)
    End Select
    DoEvents
    argContent(i) = Trim(Mid(parsestring, subPos, (argPos(i) - subPos)))
Next i
'end get argument content

Parse = argContent(argNum)
End Function


   Public Function GetFromINI(Section As String, ByVal Key As String, Directory As String) As String
   'Call WriteToINI("Header", "Color", "Black", "C:/Windows/whatever
   'Stuff$ = GetFromINI("Header", "Color", "C:/Windows/whatever.ini")
       Dim strBuffer As String
       strBuffer = String(750, Chr(0))
       Key$ = LCase$(Key$)
       GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
   End Function

   Public Sub WriteToINI(Section As String, ByVal Key As String, ByVal KeyValue As String, Directory As String)
       Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
   End Sub

Public Sub getSettings()
    d.winamp_path = GetFromINI("main", "winamp_path", sfile)
    d.mdload_delay = Val(GetFromINI("main", "mdload_delay", sfile))
End Sub
Public Sub setSettings()
    WriteToINI "main", "mdload_delay", d.winamp_path, sfile
    WriteToINI "main", "mdload_delay", d.winamp_path, sfile
End Sub




Public Function isWinampRunning(myName As String) As Boolean
Dim ak As Boolean
ak = True

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    Dim K As Integer
    Dim zz As Boolean
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        'List1.AddItem (szExename), K
        If szExename = myName Then
            zz = True
        End If
        K = K + 1

        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop

    Call CloseHandle(hSnapshot)

isWinampRunning = zz
Finish:
End Function

