Attribute VB_Name = "MyModule"
' MyModule has the following functions/subs
'    Subs:
'   OpenApp(FileName as string)
'           * Opens an application with the filename specified
'           * Use: OpenApp(FileName)
'   DisableCtrlAltDelete(bDisabled As Boolean)
'           * Disables Control Alt Delete Breaking as well as Ctrl-Escape
'           * Use: Call DisableCtrlAltDelete(Boolean)
'   Center(FormName as form)
'           * Centers the specified form
'           * Use: Center(FormName)
'   AlwaysOnTop(FormName as Form, bOnTop as boolean)
'           * Sets the named form as always on top
'           * Use: Call AlwaysOnTop(True)
'   ExitWindows(Mode)
'           * Shutdown or reboots windows
'           * Use: ExitWindows(shutdown)
'           *      ExitWindows(reboot)


' Used for DisableCtrlAltDelete
Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long
'-------------------------------------------

' Used for ExitWindows
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Private Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved _
As Long) As Long
' ---------------------

' Used for AlwaysOnTop
Const FLAGS = 3
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Public SetTop As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
'-------------------------------------------

Sub ExitWindows(ExitMode As String)
 Select Case ExitMode
 Case Is = shutdown
     t& = ExitWindowsEx(EWX_SHUTDOWN, 0)
 Case Is = reboot
     t& = ExitWindowsEx(EWX_REBOOT Or EXW_FORCE, 0)
 Case Else
    MsgBox ("Error in ExitWindows call")
 End Select
  
 End Sub
Sub AlwaysOnTop(FormName As Form, bOnTop As Boolean)
    'Sets a form as always on top
Dim Success As Integer
If bOnTop = False Then
    Success% = SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
Else
    Success% = SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End If
End Sub




Sub Center(FormName As Form)
 ' Center Forms...
 Move (Screen.Width - FormName.Width) \ 2, (Screen.Height - FormName.Height) \ 2
End Sub
Sub DisableCtrlAltDelete(bDisabled As Boolean)
    ' Disables Control Alt Delete Breaking as well as Ctrl-Escape
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

Sub OpenApp(File As String)
    'Shells to another application
    X = Shell(File)
End Sub
