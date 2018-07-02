Attribute VB_Name = "SysTray"
'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type


Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA

Public Sub AddTrayIcon(Icon As Long, Form As Object, Optional ToolTip As String)
   'Click this button to add an icon to the taskbar status area.

   'Set the individual values of the NOTIFYICONDATA data type.
   nid.cbSize = Len(nid)
   nid.hwnd = Form.hwnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = Icon
   nid.szTip = ToolTip & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid
End Sub

Public Sub RemoveTrayIcon()
   'Click this button to delete the added icon from the taskbar
   'status area by calling the Shell_NotifyIcon function.
   Shell_NotifyIcon NIM_DELETE, nid
End Sub


Public Property Get TrayEvent(mouseX As Single) As String
   
   Dim Msg As Long
   Dim sFilter As String
   Msg = mouseX / Screen.TwipsPerPixelX
    Select Case Msg
       Case WM_LBUTTONDOWN
TrayEvent = "LEFTDOWN"
       Case WM_LBUTTONUP
TrayEvent = "LEFTUP"
       Case WM_LBUTTONDBLCLK
TrayEvent = "LEFTDOUBLE"
       Case WM_RBUTTONDOWN
TrayEvent = "RIGHTDOWN"
       Case WM_RBUTTONUP
TrayEvent = "RIGHTUP"
       Case WM_RBUTTONDBLCLK
TrayEvent = "RIGHTDOUBLE"
    End Select
End Property

Public Sub TrayToolTip(Message As String)

   nid.szTip = Message & vbNullChar


   Shell_NotifyIcon NIM_MODIFY, nid
End Sub


Public Sub ChangeTrayIcon(Icon As Long)

   nid.hIcon = Icon

   Shell_NotifyIcon NIM_MODIFY, nid
End Sub


