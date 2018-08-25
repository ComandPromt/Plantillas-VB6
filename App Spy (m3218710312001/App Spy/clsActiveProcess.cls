VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Dim ListOfActiveProcess() As PROCESSENTRY32
Public Function GetActiveProcess() As Long
Dim hSnapshot As Long
Dim tProcess As PROCESSENTRY32
Dim R As Long, i As Integer
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapshot = 0 Then
    GetActiveProcess = 0
    Exit Function
End If
With tProcess
    .dwSize = Len(tProcess)
    R = ProcessFirst(hSnapshot, tProcess)
    ReDim Preserve ListOfActiveProcess(20)
    Do While R
        i = i + 1
    If i Mod 20 = 0 Then ReDim Preserve ListOfActiveProcess(i + 20)
        ListOfActiveProcess(i) = tProcess
        R = ProcessNext(hSnapshot, tProcess)
    Loop
End With
GetActiveProcess = i
Call CloseHandle(hSnapshot)
End Function
Public Function exePath(ByVal Index As Long) As String
exePath = ListOfActiveProcess(Index).szExeFile
End Function
Public Function Usage(ByVal Index As Long) As Long
Usage = ListOfActiveProcess(Index).cntUsage
End Function
Public Function Threads(ByVal Index As Long) As Long
Threads = ListOfActiveProcess(Index).cntThreads
End Function
Public Function PriClassBase(ByVal Index As Long) As Long
PriClassBase = ListOfActiveProcess(Index).pcPriClassBase
End Function
Public Function ParentProcessID(ByVal Index As Long) As Long
ParentProcessID = ListOfActiveProcess(Index).th32ParentProcessID
End Function
Public Function ProcessID(ByVal Index As Long) As Long
ProcessID = ListOfActiveProcess(Index).th32ProcessID
End Function
Public Function ModuleID(ByVal Index As Long) As Long
ModuleID = ListOfActiveProcess(Index).th32ModuleID
End Function
Public Function DefaultHeapID(ByVal Index As Long) As Long
DefaultHeapID = ListOfActiveProcess(Index).th32DefaultHeapID
End Function
Public Function dwSize(ByVal Index As Long) As Long
dwSize = ListOfActiveProcess(Index).dwSize
End Function
Public Function dwFlags(ByVal Index As Long) As Long
dwFlags = ListOfActiveProcess(Index).dwFlags
End Function
