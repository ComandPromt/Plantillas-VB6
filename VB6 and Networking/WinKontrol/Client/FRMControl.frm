VERSION 5.00
Begin VB.Form FRMControl 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ControlTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   360
   End
   Begin VB.Image ServerScreen 
      Height          =   2415
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "FRMControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ControlTimer_Timer()

Dim pt As POINTAPI
Dim rv As Long
DontCount = DontCount + 1

rv& = GetCursorPos(pt) 'Read cursor location
ReDim Preserve cp(idx) ' Create Array Members

If DontCount > 50 Then 'WAIT! Dont count First Keystroke
    If GetAsyncKeyState(VK_LBUTTON) Then
        'Label1.Caption = "Left Button Down"
        ButtonDown = "L"
    ElseIf GetAsyncKeyState(VK_RBUTTON) Then
        'Label1.Caption = "Right Button Down"
        ButtonDown = "R"
    Else
        'Label1.Caption = ""
        ButtonDown = "0"
    End If
End If

If pt.x = lastX And pt.y = LastY And ButtonDown = LastMouseButton Then
        'Dont Send Data to client
    Else
        'Image2.Visible = False
        'LightCounter.Enabled = True
        SendData "POS," & Int(pt.x / XModifier) & ":" & Int(pt.y / YModifier) & ":" & ButtonDown
End If

lastX = pt.x
LastY = pt.y
LastMouseButton = ButtonDown




End Sub

Private Sub Form_Load()

'Set The Option Panel Up
FRMControloptions.Top = Screen.Height - 1200
FRMControloptions.Left = Screen.Width - 5000
FRMControloptions.Show , FRMControl

'Setup Server Screen Image display
FRMControl.ServerScreen.Top = 1: FRMControl.ServerScreen.Left = 1
FRMControl.ServerScreen.Width = Screen.Width
FRMControl.ServerScreen.Height = Screen.Height

'Start Recording/ Sending Mouse Commands
ControlTimer.Interval = DefaultTimerInterval
ControlTimer.Enabled = True

SendData "SCREENSHOT,"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload FRMControloptions

End Sub

