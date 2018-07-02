VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormScan2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crack80"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmScan2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox checkTrace 
      BackColor       =   &H00000000&
      Caption         =   "Trace"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   5040
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin MSWinsockLib.Winsock tcpScan 
      Index           =   0
      Left            =   480
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox ListIP 
      BackColor       =   &H00C0C0C0&
      Height          =   3210
      ItemData        =   "FrmScan2.frx":0442
      Left            =   120
      List            =   "FrmScan2.frx":0444
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2160
      Width           =   5775
   End
   Begin VB.CommandButton CmdGo 
      BackColor       =   &H00808080&
      Caption         =   "Go"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Démarrer le scan"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Retour au menu"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TxtRemotePort 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Text            =   "139"
      ToolTipText     =   "Port de communication de début"
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame FrameParam 
      BackColor       =   &H00000000&
      Caption         =   "Scan d'une tranche d'IP"
      ForeColor       =   &H0000C000&
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox IPa1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Text            =   "254"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox IPd4 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox IPd3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox IPd2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox IPd1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "IP de fin :"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "IP de départ :"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "port :"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "En attente !"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label port 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "IP"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "FormScan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public portOK As Boolean

Private Sub Form_Activate()
    Call couleur(Me)
End Sub

Private Sub CmdQuit_Click()
    FormScan2.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub

Private Sub Form_Load()
    IPd1.Text = IP1
    IPd2.Text = IP2
    IPd3.Text = IP3
End Sub

Private Sub IPd1_GotFocus()
    IPd1.SelStart = 0
    IPd1.SelLength = Len(IPd1.Text)
End Sub
Private Sub IPd2_GotFocus()
    IPd2.SelStart = 0
    IPd2.SelLength = Len(IPd2.Text)
End Sub
Private Sub IPd3_GotFocus()
    IPd3.SelStart = 0
    IPd3.SelLength = Len(IPd3.Text)
End Sub
Private Sub IPd4_GotFocus()
    IPd4.SelStart = 0
    IPd4.SelLength = Len(IPd4.Text)
End Sub
Private Sub IPa1_GotFocus()
    IPa1.SelStart = 0
    IPa1.SelLength = Len(IPa1.Text)
End Sub

Private Sub TxtRemotePortD_GotFocus()
    TxtRemotePortD.SelStart = 0
    TxtRemotePortD.SelLength = Len(TxtRemotePortD.Text)
End Sub

Private Sub TxtRemotePort_LostFocus()
    If Val(TxtRemotePort.Text) > 65535 Then
        TxtRemotePort.Text = ""
        TxtRemotePort.SetFocus
    End If
End Sub

Private Sub TxtRemotePort_GotFocus()
    TxtRemotePort.SelStart = 0
    TxtRemotePort.SelLength = Len(TxtRemotePort.Text)
End Sub

Private Sub cmdGo_Click()
Dim Socket As Variant
Dim CurrentPort As Integer
Const MaxSockets = 100 'entre 1 et 200
Dim IPScanne As Integer

cpt1 = 0
cpt2 = 0
IPScanne = Val(IPd4)
On Error Resume Next

cmdQuit.Enabled = False
If cmdGo.Caption = "Go" Then
    ListIP.AddItem ("Scan du port : " & TxtRemotePort.Text)
    cmdGo.Caption = "Stop"
    Label5.Caption = "Initialisation"
    For i = 1 To MaxSockets
        'creation des instances de sock
        Load tcpScan(i)
    Next i
    Label5.Caption = "Scan en cours"
    While cmdGo.Caption = "Stop"
        For Each Socket In tcpScan
            DoEvents
            If Socket.State <> sckClosed Then
                GoTo continue
            End If
            Socket.Close
            If IPScanne = Val(IPa1) + 1 _
                Then Exit For
            Socket.RemoteHost = IPd1 & "." & IPd2 & "." & IPd3 & "." & Mid(Str(IPScanne), 2, Len(Str(IPScanne)))
            Socket.RemotePort = TxtRemotePort
            port.Caption = Socket.RemoteHost
            cpt1 = cpt1 + 1
            Socket.Connect
            IPScanne = IPScanne + 1

continue:
        Next Socket
    DoEvents
    Wend
cmdGo.Caption = "Go"
Else
    cmdGo.Caption = "Go"
End If
    ' fermeture des sockets
Label5.Caption = "Effacement mémoire"
For i = 1 To MaxSockets
    Unload tcpScan(i)
Next i
cmdQuit.Enabled = True
Label5.Caption = "En attente !"
port.Caption = ""
End Sub

Private Sub tcpScan_Connect(index As Integer)
    ListIP.AddItem ("IP OK : " & tcpScan(index).RemoteHost & " " & DottedIPToDNS(tcpScan(index).RemoteHost))
    tcpScan(index).Close
    cpt2 = cpt2 + 1
    Label4.Caption = Str(cpt2) & " IP trouvées"
    If checkTrace.Value = 1 Then
        Call ecrire_log("Port " & TxtRemotePort.Text & " OK : " & tcpScan(index).RemoteHost & " " & DottedIPToDNS(tcpScan(index).RemoteHost))
    End If
    If cpt1 = cpt2 Then
        cmdGo.Caption = "Go"
    End If
End Sub

Private Sub tcpScan_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    tcpScan(index).Close
    cpt2 = cpt2 + 1
    If cpt1 = cpt2 Then
        cmdGo.Caption = "Go"
    End If
End Sub
