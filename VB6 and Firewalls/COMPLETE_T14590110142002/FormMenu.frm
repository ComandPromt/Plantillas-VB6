VERSION 5.00
Begin VB.Form FormMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crack80"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6600
   Begin VB.CommandButton CmdPwd 
      BackColor       =   &H00808080&
      Caption         =   "Pwd"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "filtre un port en entrée"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdPort2 
      BackColor       =   &H00808080&
      Caption         =   "&Port2"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Scanne d'un port pour une tranche d'IP"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdFirewall 
      BackColor       =   &H00808080&
      Caption         =   "Firewall"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "filtre un port en entrée"
      Top             =   480
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2040
      Top             =   240
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "&Quit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Terminer Crack80 !"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdWhois 
      BackColor       =   &H00808080&
      Caption         =   "&Whois"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Whois... qui est qui ?"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdFinger 
      BackColor       =   &H00808080&
      Caption         =   "&Finger"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Finger sur un serveur"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdMail 
      BackColor       =   &H00808080&
      Caption         =   "&Mail"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Permet d'envoyer un E-mail anonyme"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdPing 
      BackColor       =   &H00808080&
      Caption         =   "Ping"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "ping une IP"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdWeb 
      BackColor       =   &H00808080&
      Caption         =   "&CGI"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Recherche les failles CGI d'un serveur"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdPort 
      BackColor       =   &H00808080&
      Caption         =   "Port "
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Scanne les ports d'une machine"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdBlaireau 
      BackColor       =   &H00808080&
      Caption         =   "&WinShare"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "scanne les ressources partagées pour une tranche d'IP... et permet de s'y connecter"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdIRC 
      BackColor       =   &H00808080&
      Caption         =   "IRC"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "connexion à un serveur IRC"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "sibair"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      ToolTipText     =   "programme écrit par sibair"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "crack80"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      ToolTipText     =   "programme écrit par sibair"
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBlaireau_Click()
    Me.Visible = False
    Load FormBlaireau
    FormBlaireau.Visible = True
    FormBlaireau.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub cmdFirewall_Click()
    Me.Visible = False
    Load FormFirewall
    FormFirewall.Visible = True
    FormFirewall.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub CmdMail_Click()
    Me.Visible = False
    Load FormMail
    FormMail.Visible = True
    FormMail.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

'Private Sub CmdNuke_Click()
'    Me.Visible = False
'    Load FormNuke
'    FormNuke.Visible = True
'    FormNuke.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
'End Sub

Private Sub CmdPing_Click()
    Me.Visible = False
    Load FormPing
    FormPing.Visible = True
    FormPing.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub CmdPort_Click()
    Me.Visible = False
    Load FormScan
    FormScan.Visible = True
    FormScan.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub cmdPort2_Click()
    Me.Visible = False
    Load FormScan2
    FormScan2.Visible = True
    FormScan2.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub CmdPwd_Click()
    Me.Visible = False
    Load FormPwd
    FormPwd.Visible = True
    FormPwd.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub CmdQuit_Click()
    Unload Me
    End
End Sub

Private Sub CmdWeb_Click()
    Me.Visible = False
    Load FormCrack
    FormCrack.Visible = True
    FormCrack.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub CmdWhois_Click()
    Me.Visible = False
    Load FormWhois
    FormWhois.Visible = True
    FormWhois.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub CmdFinger_Click()
    Me.Visible = False
    Load FormFinger
    FormFinger.Visible = True
    FormFinger.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub CmdIrc_Click()
    Me.Visible = False
    Load FormIrc
    FormIrc.Visible = True
    FormIrc.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)
End Sub

Private Sub Form_Activate()
    ' Degrade du fond
    Call couleur(Me)
    Timer1.Enabled = True
End Sub

Private Sub Form_Deactivate()
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Dim IP_temp As String
Dim pos As Long

'If App.PrevInstance = True Then
'    MsgBox ("Application déjà lancée !")
'    End
'End If

Call ouvrir_log
Call ecrire_log("########## Début crack80 : " & Now)

If Not WinsockInit Then
    Me.MousePointer = 0
    MsgBox "Erreur Initialisation Winsock !", vbCritical
    WSACleanup
    End
End If
HostName = GetIPHostName()
IPname = GetIPAddress()
Call GetCurrentUser

pos = InStr(IPname, ".")
IP1 = Mid(IPname, 1, pos - 1)
IP_temp = Mid(IPname, pos + 1, Len(IPname))

pos = InStr(IP_temp, ".")
IP2 = Mid(IP_temp, 1, pos - 1)
IP_temp = Mid(IP_temp, pos + 1, Len(IP_temp))

pos = InStr(IP_temp, ".")
IP3 = Mid(IP_temp, 1, pos - 1)

FormMenu.Caption = HostName & " - " & IPname & " - " & Utilisateur & " - " & EthernetAddress(0)

End Sub

Private Sub Form_Unload(Cancel As Integer)
WSACleanup
Call ecrire_log("########## Fin crack80 : " & Now)
Call fermer_log
End Sub

'---- Timer : affiche un rectangle
Private Sub Timer1_Timer()
    Label1.Left = Rnd * 330
    Label1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Label2.Left = Rnd * 330
    Label2.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub
