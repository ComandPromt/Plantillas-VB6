VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form FRMMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinKontrol Server [Version]"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FRMMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ScreenShot 
      Height          =   735
      Left            =   3600
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   120
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   10666
   End
   Begin VB.Frame StatusFrame 
      Caption         =   "Status"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Status 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "FRMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OverTime As Integer

Private Sub Form_Load()

'Set Form Title
Me.Caption = "WinKontrol Server v" & GetVersion

Sock.Listen 'Start Server & Update Status
UpdateStatus "Wink Server Started: " & Sock.LocalIP
Inconnection = False





End Sub

Private Sub Sock_Close()
    If Sock.State <> sckClosed Then Sock.Close
    Form_Load ' Resume Listening
End Sub

Private Sub Sock_ConnectionRequest(ByVal requestID As Long)

    On Error GoTo ErrorKontrol
    If Sock.State <> sckClosed Then Sock.Close  'Close SOCK if required
    Sock.Accept requestID                       'Accept Request

    Inconnection = True
    UpdateStatus "Client Connected"
    SendData "CONNECTED,"
    Exit Sub
     
ErrorKontrol:
     MsgBox Err.Description, vbCritical
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)


    Dim Command      As String
    Dim GetTheData As String
    Dim NewArrival   As String
    Dim Data         As String

    Sock.GetData GetTheData, vbString
    
    SplitInData = Split(GetTheData, vbCrLf)
    For i = 0 To UBound(SplitInData)
    NewArrival$ = SplitInData(i)
    
    Command = EvalData(NewArrival$, 1) 'Get Command from New Data (before the ,)
    Data$ = EvalData(NewArrival$, 2)   ' Get Data from new Data (After the ,)
    
    Select Case Command$
    
        'Hostname Requested... Send
        Case "HOSTNAME"
            SendData "HOSTNAME," & Sock.LocalHostName
            
        'Resolution Requested... Send
        Case "RESOLUTION"
            SendData "RESOLUTION," & Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY
            
        'Incoming Mouse Locations and Button Events
        Case "POS"
            'Split Mouse Positions into X, Y, & MouseButton
            SplitData = Split(Data$, ":")
            x = SplitData(0): y = SplitData(1): Button = SplitData(2)
            'Set MouseCursor Position
            SetCursorPos x, y
            'Read MouseButton Configs & Activate Button Accordingly
            Select Case Button
            Case "L" 'Left mouse button
                If LastMouseButton <> "L" Then Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                LastMouseButton = "L"
            Case "R" 'Right Mouse Button
                If LastMouseButton <> "R" Then Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                LastMouseButton = "R"
            Case "0" 'No Mouse Buttons
                If LastMouseButton = "L" Then
                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                ElseIf LastMouseButton = "R" Then
                    Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                End If
                LastMouseButton = "0"
            End Select
            
        'ScreenShot Requested... Send
        Case "SCREENSHOT"
            Dim c As New cDIBSection 'Get Picture space ready
            'Take screenshot of the ENTIRE screen below
            ScreenShot.Cls
            Set ScreenShot.Picture = CaptureWindow(0, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
            c.CreateFromPicture ScreenShot.Picture '"Import" screenshot to Picture Space
            SaveJPG c, DefaultScreenShotName, DefaultScreenShotSize 'Save Picture as JPEG image
            SendScreenShot DefaultScreenShotName

    
    End Select
    Next
    

End Sub

Private Sub Timer1_Timer()
OverTime = OverTime + 1

End Sub
