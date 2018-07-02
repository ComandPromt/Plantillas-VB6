Attribute VB_Name = "basVersion32"
' Module Name               basVersion
' Module Description        Module which contains ALL Declarations
'                           for API calls in the basVersion32 Library
'
Option Explicit

Declare Function GetFileVersionInfo _
    Lib "version.dll" _
    Alias "GetFileVersionInfoA" ( _
        ByVal lptstrFilename As String, _
        ByVal dwHandle As Long, ByVal dwLen As Long, _
        lpData As Any) _
    As Long

Declare Function GetFileVersionInfoSize _
    Lib "version.dll" _
    Alias "GetFileVersionInfoSizeA" ( _
        ByVal lptstrFilename As String, _
        lpdwHandle As Long) _
    As Long

Declare Function VerFindFile _
    Lib "version.dll" _
    Alias "VerFindFileA" ( _
        ByVal uFlags As Long, ByVal szFileName As String, _
        ByVal szWinDir As String, ByVal szAppDir As String, ByVal szCurDir As String, _
        lpuCurDirLen As Long, ByVal szDestDir As String, _
        lpuDestDirLen As Long) _
    As Long

Declare Function VerInstallFile _
    Lib "version.dll" _
    Alias "VerInstallFileA" ( _
        ByVal uFlags As Long, ByVal szSrcFileName As String, _
        ByVal szDestFileName As String, ByVal szSrcDir As String, _
        ByVal szDestDir As String, ByVal szCurDir As String, _
        ByVal szTmpFile As String, lpuTmpFileLen As Long) _
    As Long

