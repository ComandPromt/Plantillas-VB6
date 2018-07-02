VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormFirewall 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock sockFire 
      Left            =   1800
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00808080&
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox lstIP 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "FormFirewall.frx":0000
      Left            =   120
      List            =   "FormFirewall.frx":0002
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Text            =   "20034"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Port à filtrer :"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FormFirewall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGo_Click()
    sockFire.LocalPort = Val(txtPort.Text)
    sockFire.Listen
    cmdGo.Enabled = False
End Sub
Private Sub CmdQuit_Click()
    FormFirewall.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub
Private Sub Form_Activate()
   Call couleur(Me)
End Sub
Private Sub txtPort_gotfocus()
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
End Sub
Private Sub sockFire_ConnectionRequest(ByVal requestID As Long)
    lstIP.AddItem sockFire.RemoteHostIP
End Sub
