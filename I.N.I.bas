Attribute VB_Name = "ModuleINI"
Public Declare Function GetPrivateProfileSection Lib "kernel32" _
    Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
   As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileSection Lib "kernel32" _
    Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
    ByVal lpString As String, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
    
    Public gstrKeyValue As String * 256


