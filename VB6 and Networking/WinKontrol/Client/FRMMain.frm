VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form FRMMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinKontrol  [Version]"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FRMMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame HostDetailsFrame 
      Caption         =   "Host Details"
      Height          =   1215
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   4695
      Begin VB.CommandButton OptionsButton 
         Caption         =   "Options"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton ControlButton 
         Caption         =   "Control"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Unknown"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Unknown"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Unknown"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Mouse Modifier:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Screen Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Hostname:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame ConnectionFrame 
      Caption         =   "Connection Details"
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   4695
      Begin VB.CommandButton DisconnectButton 
         Caption         =   "Disconnect"
         Height          =   285
         Left            =   2880
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton ConnectButton 
         Caption         =   "Connect"
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox PortBox 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox IPBox 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "IP Address:"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   4200
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
            Size            =   9
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
Private Sub ConnectButton_Click()
    
    'Attempt Connection
    Replied = False
    Sock.Connect IPBox.Text, PortBox.Text
    UpdateStatus "Attempting Connection..."
    WaitTime = 0
    
    While (Not Replied) And (WaitTime < 100000)
        DoEvents
        WaitTime = WaitTime + 1
    Wend
    
    If WaitTime >= 100000 Then
        MsgBox "Unable to connect to remote server", vbCritical, "Connection Error"
        UpdateStatus "Connection Failed"
        Sock.Close
        Exit Sub
    Else
        'Enable/ Disable Certain Features
        IPBox.Enabled = False
        PortBox.Enabled = False
        ConnectButton.Enabled = False
        DisconnectButton.Enabled = True
        ControlButton.Enabled = True
        OptionsButton.Enabled = True
     End If
    
End Sub

Private Sub ControlButton_Click()

FRMControl.Top = 1: FRMControl.Left = 1
FRMControl.Width = Screen.Width: FRMControl.Height = Screen.Height
FRMControl.Show


End Sub

Private Sub DisconnectButton_Click()
    Sock.Close
    UpdateStatus "Disconnected From Server"
    Form_Load
End Sub

Private Sub Form_Load()

'Set Form Title
Me.Caption = "WinKontrol v" & GetVersion

'SetDefaultSettings
IPBox.Text = DefaultIPAddress
PortBox.Text = DefaultPort

'Activate/ Deactivate Buttons
IPBox.Enabled = True
PortBox.Enabled = True
ConnectButton.Enabled = True
DisconnectButton.Enabled = False
ControlButton.Enabled = False
OptionsButton.Enabled = False

'Clear all Previous Connection Details
Label6.Caption = "Unknown"
Label7.Caption = "Unknown"
Label8.Caption = "Unknown"





UpdateStatus "WinKontrol Sucessfully Initialised"

End Sub


Private Sub HostDetailsLabel_Click(Index As Integer)

End Sub

Private Sub Sock_Close()
    If Sock.State <> sckClosed Then Sock.Close
    Form_Load
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)

    Dim Command      As String
    Dim NewArrival   As String
    Dim Data         As String

    Sock.GetData NewArrival$, vbString
    Command = EvalData(NewArrival$, 1) 'Get Command from New Data (before the ,)
    Data$ = EvalData(NewArrival$, 2)   ' Get Data from new Data (After the ,)
    
    Select Case Command$
    
        Case "CONNECTED"
            'Get Servers conection response and request HOSTNAME
            Replied = True
            UpdateStatus "Connection Established"
            SendData "HOSTNAME,"
            
        Case "HOSTNAME"
            'Get Servers HOSTNAME and request RESOLUTION
            Label6.Caption = Data$
            SendData "RESOLUTION,"
        
        Case "RESOLUTION"
            'Get Servers resolution and calculate Mouse Modifiers
            Label7.Caption = Data$
            SplitData = Split(Data$, "x") ' Splits X/Y Into Arrays of Splitdata
            'urwidth = Screen.Width / Screen.TwipsPerPixelX
            'urheight = Screen.Height / Screen.TwipsPerPixelY
            XModifier = ((Screen.Width / Screen.TwipsPerPixelX) / SplitData(0))
            YModifier = ((Screen.Height / Screen.TwipsPerPixelY) / SplitData(1))
            Label8.Caption = XModifier & " x " & YModifier
            
        Case "SCREENSHOT"
            Close 10: Open DefaultScreenShotName For Binary Access Write As #10
            
        Case "ENDSCREENSHOT"
            ' Close File and Display screenshot
            Close 10: FRMControl.ServerScreen.Picture = LoadPicture(DefaultScreenShotName)
            Pause 200 ' Pause
            SendData "SCREENSHOT,"
        
        Case Else 'Data is therefore Screenshot DATA!
            Put #10, , NewArrival$ 'Write data to Picture file
            DoEvents
    
    End Select
    


End Sub

Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Sock.State <> sckClosed Then Sock.Close
End Sub
