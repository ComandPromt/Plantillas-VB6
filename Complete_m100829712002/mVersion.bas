Attribute VB_Name = "mVersion"
Option Explicit

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                     (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
 
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Global IsNT4 As Boolean
Global IsNT3 As Boolean
Global Is2000 As Boolean
Global Is95 As Boolean
Global Is95B As Boolean
Global Is98 As Boolean
Global Is98se As Boolean
Global IsME As Boolean

Public Sub SetWinVersion()
Dim version As OSVERSIONINFO

    version.dwOSVersionInfoSize = Len(version)
    GetVersionEx version

    If version.dwPlatformId = 1 And version.dwMinorVersion = 10 And LoWord(version.dwBuildNumber) = 1998 Then
        Is98 = True
    ElseIf version.dwPlatformId = 1 And version.dwMinorVersion = 10 And LoWord(version.dwBuildNumber) = 2222 Then
        Is98se = True
    ElseIf version.dwPlatformId = 1 And version.dwMinorVersion = 90 And LoWord(version.dwBuildNumber) = 3000 Then
        IsME = True
    ElseIf version.dwPlatformId = 1 And version.dwMinorVersion = 0 And LoWord(version.dwBuildNumber) = 950 Then
        Is95 = True
    ElseIf version.dwPlatformId = 1 And version.dwMinorVersion = 0 And LoWord(version.dwBuildNumber) = 1111 Then
        Is95B = True
    End If
            
    If version.dwPlatformId = 2 And version.dwMajorVersion = 3 Then
        IsNT3 = True
    ElseIf version.dwPlatformId = 2 And version.dwMajorVersion = 4 Then
        IsNT4 = True
    ElseIf version.dwPlatformId = 2 And version.dwMajorVersion = 5 Then
        Is2000 = True
    End If
    
End Sub

Private Function LoWord(lngIn As Long) As Integer
   If (lngIn And &HFFFF&) > &H7FFF Then
      LoWord = (lngIn And &HFFFF&) - &H10000
   Else
      LoWord = lngIn And &HFFFF&
   End If
End Function



