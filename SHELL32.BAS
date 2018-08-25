Attribute VB_Name = "basShell32"
' Module Name               basShell32
' Module Description        Module which contains ALL Declarations
'                           for API calls in the Shell32 Library
'
Option Explicit

Declare Function CommandLineToArgv _
    Lib "shell32" _
    Alias "CommandLineToArgvW" ( _
        ByVal lpCmdLine As String, _
        pNumArgs As Integer) _
    As Long

Declare Function DoEnvironmentSubst _
    Lib "shell32.dll" _
    Alias "DoEnvironmentSubstA" ( _
        ByVal szString As String, _
        ByVal cbString As Long) _
    As Long

Declare Sub DragAcceptFiles _
    Lib "shell32.dll" ( _
        ByVal hWnd As Long, _
        ByVal fAccept As Long)

Declare Sub DragFinish _
    Lib "shell32.dll" ( _
        ByVal HDROP As Long)

Declare Function DragQueryFile _
    Lib "shell32.dll" _
    Alias "DragQueryFileA" ( _
        ByVal HDROP As Long, ByVal UINT As Long, _
        ByVal lpStr As String, ByVal ch As Long) _
    As Long

Declare Function DragQueryPoint _
    Lib "shell32.dll" ( _
        ByVal HDROP As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function DuplicateIcon _
    Lib "shell32.dll" ( _
        ByVal hInst As Long, _
        ByVal hIcon As Long) _
    As Long

Declare Function ExtractAssociatedIcon _
    Lib "shell32.dll" _
    Alias "ExtractAssociateIconA" ( _
        ByVal hInst As Long, _
        ByVal lpIconPath As String, _
        lpiIcon As Long) _
    As Long

Declare Function ExtractIcon _
    Lib "shell32.dll" _
    Alias "ExtractIconA" ( _
        ByVal hInst As Long, _
        ByVal lpszExeFileName As String, _
        ByVal nIconIndex As Long) _
    As Long

Declare Function ExtractIconEx _
    Lib "shell32.dll" _
    Alias "ExtractIconExA" ( _
        ByVal lpszFile As String, ByVal nIconIndex As Long, _
        phiconLarge As Long, phiconSmall As Long, _
        ByVal nIcons As Long) _
    As Long

Declare Function FindEnvironmentString _
    Lib "shell32.dll" _
    Alias "FindEnvironmentStringA" ( _
        ByVal szEnvVar As String) _
    As String

Declare Function FindExecutable _
    Lib "shell32.dll" _
    Alias "FindExecutableA" ( _
        ByVal lpFile As String, _
        ByVal lpDirectory As String, _
        ByVal lpResult As String) _
    As Long

Declare Function SHAppBarMessage _
    Lib "shell32.dll" ( _
        ByVal dwMessage As Long, _
        pData As APPBARDATA) _
    As Long

Declare Function Shell_NotifyIcon _
    Lib "shell32.dll" _
    Alias " Shell_NotifyIconA" ( _
        ByVal dwMessage As Long, _
        lpData As NOTIFYICONDATA) _
    As Long

Declare Function ShellAbout _
    Lib "shell32.dll" _
    Alias "ShellAboutA" ( _
        ByVal hWnd As Long, ByVal szApp As String, _
        ByVal szOtherStuff As String, ByVal hIcon As Long) _
    As Long

Declare Function ShellExecute _
    Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) _
    As Long

Declare Function SHFileOperation _
    Lib "shell32.dll" _
    Alias " SHFileOperationA" ( _
        lpFileOp As SHFILEOPSTRUCT) _
    As Long

Declare Sub SHFreeNameMappings _
    Lib "shell32.dll" ( _
        ByVal hNameMappings As Long)

Declare Sub WinExecError _
    Lib "shell32.dll" _
    Alias "WinExecErrorA" ( _
        ByVal hWnd As Long, _
        ByVal error As Long, _
        ByVal lpstrFileName As String, _
        ByVal lpstrTitle As String)


