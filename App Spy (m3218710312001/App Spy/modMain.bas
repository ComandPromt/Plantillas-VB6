Attribute VB_Name = "modMain"
Option Explicit
Const MAX_MODULE_NAmeInfo = 255
Const MAX_PATH = 260
Const TH32CS_SNAPMODULE = &H8
Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * MAX_PATH
End Type
Type hKeys
  Class As Long
  Key As String
  Name As String
End Type
Public hKeys(4) As hKeys
Global MainKeys(50) As String
Global SubKeys(50) As String
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Any, ByVal source As Any, ByVal bytes As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lpmeInfo As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, lpmeInfo As MODULEENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, pdwResult As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private pid As Long
Public IsResond As String
Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lpData As Long) As Long
Dim lThreadId  As Long
Dim lProcessId As Long
fEnumWindowsCallBack = 1
lThreadId = GetWindowThreadProcessId(hwnd, lProcessId)
If lProcessId = pid Then
    Call strCheck(hwnd)
    fEnumWindowsCallBack = 0
End If
End Function
Public Function fEnumWindows(clsPID As Long) As Boolean
Dim hwnd As Long
pid = clsPID
Call EnumWindows(AddressOf fEnumWindowsCallBack, hwnd)
End Function
Private Function strCheck(ByVal lhwnd As Long)
Dim lResult As Long
Dim lReturn As Long
Dim strRunning As String
If lhwnd = 0 Then Exit Function
lReturn = SendMessageTimeout(lhwnd, &H0, 0&, 0&, &H2 And &H1, 1000, lResult)
IsResond = IIf(lReturn, "Responding", "Not Responding")
End Function
Public Function WinDir() As String
Dim buffer As String * 512, length As Integer
length = GetWindowsDirectory(buffer, Len(buffer))
WinDir = Left$(buffer, length)
End Function
Public Function ReadINI(iSection As String, iKey As String, iniFile As String)
Dim RetStr As String, Retlen As String, iPath As String
iPath = WinDir & "\" & iniFile
RetStr = Space$(255)
Retlen = GetPrivateProfileString(iSection, iKey, "", RetStr, Len(RetStr), iPath)
RetStr = Left$(RetStr, Retlen)
ReadINI = IIf(RetStr = "", "Empty...", RetStr)
End Function
Public Sub WriteINI(iSection As String, iKey As String, iniFile As String, Text As String)
WritePrivateProfileString iSection, iKey, Text, iniFile
End Sub
Function GetProcessModules(Optional ByVal ProcessID As Long = -1) As String()
Dim meInfo As MODULEENTRY32
Dim success As Long
Dim hSnapshot As Long
ReDim res(0) As String
Dim count As Long
If ProcessID = -1 Then ProcessID = GetCurrentProcessId
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPMODULE, ProcessID)
If hSnapshot = 0 Then GoTo ExitProc
meInfo.dwSize = Len(meInfo)
success = Module32First(hSnapshot, meInfo)
    Do While success
        If meInfo.th32ProcessID = ProcessID Then
            count = count + 1
            If count > UBound(res) Then ReDim Preserve res(count + 100) As String
            res(count) = Left$(meInfo.szExePath, InStr(meInfo.szExePath & vbNullChar, vbNullChar) - 1)
        End If
        success = Module32Next(hSnapshot, meInfo)
    Loop
    CloseHandle hSnapshot
ExitProc:
    ReDim Preserve res(0 To count) As String
    GetProcessModules = res
End Function
Sub lKeys()
Dim cver As String
cver = "Software\Microsoft\Windows\CurrentVersion\"
hKeys(1).Class = &H80000001
hKeys(1).Key = cver & "Run"
hKeys(1).Name = "User_Run"
hKeys(2).Class = &H80000001
hKeys(2).Key = cver & "RunServices"
hKeys(2).Name = "User_RunServices"
hKeys(3).Class = &H80000002
hKeys(3).Key = cver & "Run"
hKeys(3).Name = "Machine_Run"
hKeys(4).Class = &H80000002
hKeys(4).Key = cver & "RunServices"
hKeys(4).Name = "Machine_RunServices"
End Sub

