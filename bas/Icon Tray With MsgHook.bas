Attribute VB_Name = "MsgHook32_Bas"
Option Explicit

Type NOTIFYICONDATA
    lStructureSize    As Long
    hwnd   As Long
    lID As Long
    lFlags As Long
    lCallBackMessage As Long
    hIcon As Long
    sTip As String * 64
End Type

Type lRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type APPBARDATA
    lStructureSize As Long
    hwnd As Long
    lCallBackMessage As Long
    lEdge As Long
    rc As lRect
    lParam As Long
End Type


Declare Function Shell_NotifyIcon& Lib "shell32.dll" (ByVal lMessage&, NID As NOTIFYICONDATA)
Declare Function SHAppBarMessage& Lib "shell32.dll" (ByVal dwMessage&, pData As APPBARDATA)

Global idShell_NotifyIcon&
Global idSHAppBarMessage&

Global Const NIM_ADD = 0&
Global Const NIM_DELETE = 2&
Global Const NIM_MODIFY = 1&
Global Const NIF_ICON = 2&
Global Const NIF_MESSAGE = 1&
Global Const NIF_TIP = 4&

Global Const ABM_GETTASKBARPOS = &H5&

Global structNotify As NOTIFYICONDATA
Global structBarData As APPBARDATA

'Message blaster callback stuff
Const WM_USER = &H400
Global Const UM_TASKBARMESSAGE = WM_USER + &H201
Global Const POSTPROCESS = 1
Sub modIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
    Dim ltemplong  As Long
    structNotify.lStructureSize = 88&
    structNotify.hwnd = Form1.hwnd
    structNotify.lID = IconID
    structNotify.lFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    structNotify.lCallBackMessage = UM_TASKBARMESSAGE
    structNotify.hIcon = Icon
    structNotify.sTip = ToolTip & Chr$(0)
    ltemplong = Shell_NotifyIcon(NIM_MODIFY, structNotify)
End Sub



Sub AddIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
    structBarData.lStructureSize = 36&
    Dim ltemplong As Long
    ltemplong = SHAppBarMessage(ABM_GETTASKBARPOS, structBarData)
    If ltemplong <> 1 Then
        MsgBox "Explorer Not Running! Exiting...", 16, App.Title
        End
        Exit Sub
    End If
    structNotify.lStructureSize = 88&
    structNotify.hwnd = Form1.hwnd
    structNotify.lID = IconID
    structNotify.lFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    structNotify.lCallBackMessage = UM_TASKBARMESSAGE
    structNotify.hIcon = Icon
    structNotify.sTip = ToolTip & Chr$(0)
    ltemplong = Shell_NotifyIcon(NIM_ADD, structNotify)
    
 Form1.Msghook1.HwndHook = Form1.hwnd
 Form1.Msghook1.Message(UM_TASKBARMESSAGE) = True
      

End Sub



Sub delIcon(IconID As Long)
    Dim ltemplong As Long
    structNotify.lID = IconID
    ltemplong = Shell_NotifyIcon(NIM_DELETE, structNotify)
End Sub


