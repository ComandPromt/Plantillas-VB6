VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormScan 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crack80"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
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
   Icon            =   "FrmScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock tcpScan 
      Index           =   0
      Left            =   360
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox ListPorts 
      BackColor       =   &H00C0C0C0&
      Height          =   3660
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
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
      TabIndex        =   4
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
      TabIndex        =   5
      ToolTipText     =   "Retour au menu"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TxtRemotePortF 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Text            =   "139"
      ToolTipText     =   "Port de communication de fin"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox TxtRemotePortD 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Port de communication de début"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox TxtRemoteHost 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "warez.xprojekt.cz"
      ToolTipText     =   "Serveur"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Frame FrameParam 
      BackColor       =   &H00000000&
      Caption         =   "Scan des ports"
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "à"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "de"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Machine ou IP"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "En attente !"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label port 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Ports ouverts"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "FormScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public portOK As Boolean

Private Sub Form_Activate()
    Call couleur(Me)
End Sub

Private Sub CmdQuit_Click()
    FormScan.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub

Private Sub TxtRemoteHost_GotFocus()
    TxtRemoteHost.SelStart = 0
    TxtRemoteHost.SelLength = Len(TxtRemoteHost.Text)
End Sub

Private Sub TxtRemotePortD_GotFocus()
    TxtRemotePortD.SelStart = 0
    TxtRemotePortD.SelLength = Len(TxtRemotePortD.Text)
End Sub

Private Sub TxtRemotePortD_LostFocus()
    If Val(TxtRemotePortD.Text) > 65535 Then
        TxtRemotePortD.Text = ""
        TxtRemotePortD.SetFocus
    End If
End Sub

Private Sub TxtRemotePortF_LostFocus()
    If Val(TxtRemotePortF.Text) > 65535 Then
        TxtRemotePortF.Text = ""
        TxtRemotePortF.SetFocus
    End If
End Sub

Private Sub TxtRemotePortF_GotFocus()
    TxtRemotePortF.SelStart = 0
    TxtRemotePortF.SelLength = Len(TxtRemotePortF.Text)
End Sub

Private Sub cmdGo_Click()
Dim Socket As Variant
Dim CurrentPort As Integer
Const MaxSockets = 100 'entre 1 et 200

cpt1 = 0
cpt2 = 0

On Error Resume Next

CmdQuit.Enabled = False
If cmdGo.Caption = "Go" Then
    ListPorts.AddItem ("Scan de : " & TxtRemoteHost.Text)
    cmdGo.Caption = "Stop"
    Label5.Caption = "Initialisation"
    For i = 1 To MaxSockets
        'creation des instances de sock
        Load tcpScan(i)
    Next i
    CurrentPort = TxtRemotePortD.Text
    Label5.Caption = "Scan en cours"
    While cmdGo.Caption = "Stop"
        For Each Socket In tcpScan
            DoEvents
            If Socket.State <> sckClosed Then
                GoTo continue
            End If
            Socket.Close
            If CurrentPort = Val(TxtRemotePortF.Text) + 1 _
                Then Exit For
            Socket.RemoteHost = TxtRemoteHost.Text
            Socket.RemotePort = CurrentPort
            port.Caption = CurrentPort
            cpt1 = cpt1 + 1
            Socket.Connect
            CurrentPort = CurrentPort + 1

continue:
        Next Socket
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
CmdQuit.Enabled = True
Label5.Caption = "En attente !"
port.Caption = ""
End Sub

Private Sub tcpScan_Connect(index As Integer)
    ListPorts.AddItem ("Port OK : " & Str(tcpScan(index).RemotePort))
    tcpScan(index).Close
    cpt2 = cpt2 + 1
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
