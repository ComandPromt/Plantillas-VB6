Attribute VB_Name = "ModGen"
Option Explicit
'fileexists APIs
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
'searching combobox
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long
Public Const CB_FINDSTRING = &H14C
Public safesavename As String 'variable to hold new filename

Public Function FileExists(sSource As String) As Boolean
'checks for a file's existance
If Right(sSource, 2) = ":\" Then
    Dim allDrives As String
    allDrives = Space$(64)
    Call GetLogicalDriveStrings(Len(allDrives), allDrives)
    FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
    Exit Function
Else
    If Not sSource = "" Then
        Dim WFD As WIN32_FIND_DATA
        Dim hFile As Long
        hFile = FindFirstFile(sSource, WFD)
        FileExists = hFile <> INVALID_HANDLE_VALUE
        Call FindClose(hFile)
    Else
        FileExists = False
    End If
End If
End Function
Public Function SafeSave(Path As String) As String
    'ensures a unique file name by adding a number as appropriate
    Dim mPath As String, mname As String, mTemp As String, mfile As String, mExt As String, m As Integer
    On Error Resume Next
    mPath = mID$(Path, 1, InStrRev(Path, "\"))
    mname = mID$(Path, InStrRev(Path, "\") + 1)
    mfile = Left(mID$(mname, 1, InStrRev(mname, ".")), Len(mID$(mname, 1, InStrRev(mname, "."))) - 1)
    If mfile = "" Then mfile = mname
    mExt = mID$(mname, InStrRev(mname, "."))
    mTemp = ""
    Do
        If Not FileExists(mPath + mfile + mTemp + mExt) Then
            SafeSave = mPath + mfile + mTemp + mExt
            safesavename = mfile + mTemp + mExt
            Exit Do
        End If
        m = m + 1
        mTemp = Right(Str(m), Len(Str(m)) - 1)
    Loop
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    'simple string parse
    Dim temp As String
    If InStr(1, filepath, ".") = 0 Then
        temp = filepath
    Else
        temp = mID$(filepath, 1, InStrRev(filepath, "."))
        temp = Left(temp, Len(temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = temp + newext
End Function
Public Function PathOnly(ByVal filepath As String) As String
    'simple string parse
    Dim temp As String
    temp = mID$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Public Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    'simple string parse
    ExtOnly = mID$(filepath, InStrRev(filepath, ".") + 1)
    If dot = True Then ExtOnly = "." + ExtOnly
End Function



