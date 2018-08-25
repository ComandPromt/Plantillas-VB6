Attribute VB_Name = "Task_ProcessInfo"
'i dont use much of this modules' code in my app due to the OS compatibilities.
'i have it here just for future use of any other user who wants to use this code
'in their own app.

Option Explicit
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Public Enum GetPriority
    Low = &H40
    BelowNormal = &H4000
    Normal = &H20
    AboveNormal = &H8000
    High = &H80
    Realtime = &H100
End Enum
Public Const MAX_PATH As Long = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwflags As Long
    szexeFile As String * MAX_PATH
End Type
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlgas As Long, ByVal lProcessID As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const PROCESS_TERMINATE = &H1&
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const PROCESS_SET_INFORMATION = &H200&
Private Const PROCESS_QUERY_INFORMATION = &H400&

Public Function ConvertComboBoxValToPriority(Cbv As Long) As Long

    Select Case Cbv
      Case 0
        ConvertComboBoxValToPriority = GetPriority.Low
      Case 1
        ConvertComboBoxValToPriority = GetPriority.BelowNormal
      Case 2
        ConvertComboBoxValToPriority = GetPriority.Normal
      Case 3
        ConvertComboBoxValToPriority = GetPriority.AboveNormal 'for some reason Prior can = &h8000 but it wont respond to this..
      Case 4
        ConvertComboBoxValToPriority = GetPriority.High
      Case 5
        ConvertComboBoxValToPriority = GetPriority.Realtime
      Case Else
        ConvertComboBoxValToPriority = GetPriority.AboveNormal  'used cuz the weird event for AboveNormal not working..
    End Select

End Function

Public Function ConvertPriorityToComboBoxVal(Prior As Long) As Long

    Select Case Prior
      Case Is = GetPriority.Low
        ConvertPriorityToComboBoxVal = 0
      Case Is = GetPriority.BelowNormal
        ConvertPriorityToComboBoxVal = 1
      Case Is = GetPriority.Normal
        ConvertPriorityToComboBoxVal = 2
      Case Is = GetPriority.AboveNormal 'for some reason Prior can = &h8000 but it wont respond to this..
        ConvertPriorityToComboBoxVal = 3
      Case Is = GetPriority.High
        ConvertPriorityToComboBoxVal = 4
      Case Is = GetPriority.Realtime
        ConvertPriorityToComboBoxVal = 5
      Case Else
        ConvertPriorityToComboBoxVal = 3 'used cuz the weird event for AboveNormal not working..
    End Select

End Function

'used to close a process
Public Sub EndProcess(Process As Long)

  Dim handle As Long
  Dim ExitCode As Long

    handle = OpenProcess(PROCESS_TERMINATE, 0, Process)
    GetExitCodeProcess handle, ExitCode
    TerminateProcess handle, ExitCode
    CloseHandle handle

End Sub

Public Function Get_Thread_ProcessID(hwnd As Long, ByRef RetValProc As Long)

    Get_Thread_ProcessID = GetWindowThreadProcessId(hwnd, RetValProc)

End Function

Public Function GetExeFromHandle(wnd As Long) As String

  Dim ThreadId As Long, ProcessId As Long, hSnapshot As Long
  Dim uProcess As PROCESSENTRY32, rProcessFound As Long
  Dim i As Integer, szExename As String

    ' Get ID for window thread
    ThreadId = GetWindowThreadProcessId(wnd, ProcessId)
    ' Check if valid
    If ThreadId = 0 Or ProcessId = 0 Then
        Exit Function
    End If
    ' Create snapshot of current processes
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    ' Check if snapshot is valid
    If hSnapshot = -1 Then
        Exit Function
    End If
    'Initialize uProcess with correct size
    uProcess.dwSize = Len(uProcess)
    'Start looping through processes
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        If uProcess.th32ProcessID = ProcessId Then
            'Found it, now get name of exefile
            i = InStr(1, uProcess.szexeFile, Chr$(0))
            If i Then
                szExename = Left$(uProcess.szexeFile, i - 1)
            End If
            Exit Do
          Else
            'Wrong ID, so continue looping
            rProcessFound = ProcessNext(hSnapshot, uProcess)
        End If
    Loop
    CloseHandle hSnapshot
    GetExeFromHandle = szExename

End Function

Public Function GetProcessPriority(DaProcess As Long) As Long

  Dim RetVal As Long

    RetVal = OpenProcess(PROCESS_QUERY_INFORMATION, 0, DaProcess)
    GetProcessPriority = GetPriorityClass(RetVal)
    CloseHandle RetVal

End Function

Public Function SetProcessPriority(DaProcess As Long, NewIndex As Long) As GetPriority

  Dim RetVal As Long

    RetVal = OpenProcess(PROCESS_SET_INFORMATION, 0, DaProcess)
    SetPriorityClass RetVal, ConvertComboBoxValToPriority(NewIndex)
    CloseHandle RetVal
    SetProcessPriority = ConvertPriorityToComboBoxVal(GetProcessPriority(DaProcess))

End Function


'The below are NT only!
Public Function GetUserTime(Process As Long) As String

  Dim FTnull As FILETIME, FT As FILETIME, st As SYSTEMTIME

    GetProcessTimes Process, FTnull, FTnull, FT, FTnull
    FileTimeToLocalFileTime FT, FT
    FileTimeToSystemTime FT, st
    GetUserTime = CStr(st.wHour) + ":" + CStr(st.wMinute) + "." + CStr(st.wSecond) + " on " + CStr(st.wMonth) + "/" + CStr(st.wDay) + "/" + CStr(st.wYear)

End Function
Public Function GetKernelTime(Process As Long) As String

  Dim FTnull As FILETIME, FT As FILETIME, st As SYSTEMTIME

    GetProcessTimes Process, FTnull, FTnull, FT, FTnull
    FileTimeToLocalFileTime FT, FT
    FileTimeToSystemTime FT, st
    GetKernelTime = CStr(st.wHour) + ":" + CStr(st.wMinute) + "." + CStr(st.wSecond) + " on " + CStr(st.wMonth) + "/" + CStr(st.wDay) + "/" + CStr(st.wYear)

End Function
Public Function GetCreationTime(Process As Long) As String

  Dim FTnull As FILETIME, FT As FILETIME, st As SYSTEMTIME

    GetProcessTimes Process, FT, FTnull, FTnull, FTnull
    FileTimeToLocalFileTime FT, FT
    FileTimeToSystemTime FT, st
    GetCreationTime = CStr(st.wHour) + ":" + CStr(st.wMinute) + "." + CStr(st.wSecond) + " on " + CStr(st.wMonth) + "/" + CStr(st.wDay) + "/" + CStr(st.wYear)

End Function
Public Function GetClosingTime(Process As Long) As String

  Dim FTnull As FILETIME, FT As FILETIME, st As SYSTEMTIME

    GetProcessTimes Process, FTnull, FT, FTnull, FTnull
    FileTimeToLocalFileTime FT, FT
    FileTimeToSystemTime FT, st
    GetClosingTime = CStr(st.wHour) + ":" + CStr(st.wMinute) + "." + CStr(st.wSecond) + " on " + CStr(st.wMonth) + "/" + CStr(st.wDay) + "/" + CStr(st.wYear)

End Function
