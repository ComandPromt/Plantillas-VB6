VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormCrack 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crack80"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   ClipControls    =   0   'False
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
   Icon            =   "FrmCrack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdEnvoi 
      BackColor       =   &H00808080&
      Caption         =   "Envoi"
      Enabled         =   0   'False
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
      Left            =   3240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Déconnexion du serveur"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame FramePossible 
      BackColor       =   &H00000000&
      Caption         =   "Failles"
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   6360
      TabIndex        =   17
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton OptCertaine 
         BackColor       =   &H00000000&
         Caption         =   "Certaines"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptPossible 
         BackColor       =   &H00000000&
         Caption         =   "Possibles"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
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
      Left            =   6360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Retour au menu"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton CmdDeconnect 
      BackColor       =   &H00808080&
      Caption         =   "Déconnecte"
      Enabled         =   0   'False
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
      Left            =   1680
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Déconnexion du serveur"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H00808080&
      Caption         =   "Connecte"
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
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Connexion au serveur"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox TxtRemotePort 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "80"
      ToolTipText     =   "Port de communication"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TxtRemoteHost 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "www.oreka.fr"
      ToolTipText     =   "Serveur"
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox TxtOutput 
      BackColor       =   &H00C0C0C0&
      Height          =   4575
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Faille à tester"
      Top             =   600
      Width           =   6135
   End
   Begin VB.CommandButton CmdTot 
      BackColor       =   &H00808080&
      Caption         =   "Total"
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
      Left            =   4800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Recherche des failles contenues dans Crack80.dat"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox ListOK 
      BackColor       =   &H00C0C0C0&
      Height          =   4560
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   5760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Lblstatut 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Failles détectées"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Failles testées"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label compte 
      BackColor       =   &H00808080&
      Caption         =   "0"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label txtcompteF 
      BackColor       =   &H00808080&
      Caption         =   "0"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "FormCrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public recu As Boolean
Public ligne As String
Public total As Boolean
Public compteur As Integer
Public compteF As Integer
Private Sub Form_Activate()
    Call couleur(Me)
End Sub
Private Sub CmdDeconnect_Click()

tcpClient.Close
cmdConnect.Enabled = True
CmdEnvoi.Enabled = False
CmdDeconnect.Enabled = False
TxtOutput = ""
txtSend = ""
Lblstatut.Caption = ""

End Sub

Private Sub CmdEnvoi_Click()
Dim txtTemp
    total = False
    txtTemp = txtSend.Text
    txtSend.Text = txtSend.Text & vbCrLf & vbCrLf
    tcpClient.SendData txtSend.Text
    txtSend.Text = txtTemp
    txtSend.SetFocus
    txtSend.SelStart = 0
    txtSend.SelLength = Len(txtSend.Text)
        
End Sub

Private Sub CmdQuit_Click()
    FormCrack.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub

Private Sub CmdTot_Click()
Dim num_fichier As Integer
Dim texte_envoye As String

compteur = 0
compteF = 0
total = True
TxtOutput.Text = ""
ListOK.Clear
On Error GoTo Err_Fichier
FormCrack.MousePointer = vbHourglass

tcpClient.Close

tcpClient.RemoteHost = TxtRemoteHost.Text
tcpClient.RemotePort = Val(TxtRemotePort)

num_fichier = FreeFile
Open ".\Crack80.dat" For Input As #num_fichier
'
' Attention aux erreurs
        On Error Resume Next
'
Do
    Line Input #num_fichier, ligne
    compteur = compteur + 1
    tcpClient.Connect
    While tcpClient.State <> 7
        If tcpClient.State = 9 Then
            MsgBox ("Le serveur n'est pas en service actuellement")
            End
        End If
        Lblstatut.Caption = "statut : " & Str(tcpClient.State)
        DoEvents
    Wend
    compte.Caption = Str(compteur)
    TxtOutput.Text = TxtOutput.Text & ligne & vbCrLf
    texte_envoye = ligne & vbCrLf & vbCrLf
    txtSend.Text = ligne
    tcpClient.SendData texte_envoye
    recu = False
    
    ' Attente de reception des donnees
    While recu = False
        DoEvents
    Wend

Loop Until EOF(num_fichier)

'cmdConnect.Enabled = True

Close #num_fichier

FormCrack.MousePointer = vbDefault
Exit Sub

Err_Fichier:
    MsgBox ("Erreur sur le fichier Crack80.dat")
End Sub

Private Sub Form_Load()

tcpClient.RemoteHost = TxtRemoteHost
tcpClient.RemotePort = TxtRemotePort
CmdDeconnect.Enabled = False
CmdEnvoi.Enabled = False
OptCertaine = True

End Sub

Private Sub cmdConnect_Click()

FormCrack.MousePointer = vbHourglass

compteur = 0
compteF = 0

tcpClient.RemoteHost = TxtRemoteHost
tcpClient.RemotePort = Val(TxtRemotePort)
TxtOutput.Text = ""

' Invoque la méthode Connect pour établir une connexion

' Valeurs des statuts retournes apres une demande de connexion
'sckClosed 0 Fermé (valeur par défaut)
'sckOpen 1 Ouvert
'sckListening 2 À l'écoute
'sckConnectionPending  3 Connexion en attente
'sckResolvingHost  4 Hôte en cours de résolution
'sckHostResolved  5 Hôte résolu
'sckConnecting  6 En cours de connexion
'sckConnected  7 Connecté
'sckClosing  8 Connexion en cours de fermeture par l'homologue
'sckError  9 Erreur

Lblstatut.Caption = "statut : " & Str(tcpClient.State)
tcpClient.Connect
DoEvents

While tcpClient.State <> 7
    Lblstatut.Caption = "statut : " & Str(tcpClient.State)
    DoEvents
    If tcpClient.State = 9 Then
        MsgBox ("Le serveur demandé n'est pas en service !")
        End
    End If
Wend

If tcpClient.State <> 9 Then
    cmdConnect.Enabled = False
    CmdDeconnect.Enabled = True
    CmdEnvoi.Enabled = True
    txtSend.SetFocus
Else
    MsgBox ("Le serveur demandé n'est pas en service !")
    End
End If

Lblstatut.Caption = "statut : " & Str(tcpClient.State)
cmdConnect.Enabled = False
CmdDeconnect.Enabled = True
CmdEnvoi.Enabled = True
FormCrack.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
If tcpClient.State <> 0 Then
    tcpClient.Close
End If
End Sub

Private Sub ListOK_DblClick()
txtSend.Text = ListOK.Text
TxtOutput.Text = ""
End Sub

Private Sub tcpClient_Close()

    cmdConnect.Enabled = True
    CmdDeconnect.Enabled = False
    If total = False Then
        TxtOutput.Text = TxtOutput.Text & vbCrLf & "La connexion a été interrompue par le serveur" & vbCrLf
        recu = True
    End If
    TxtOutput.SelStart = Len(TxtOutput.Text)
    
    FormCrack.MousePointer = vbHourglass
    tcpClient.Close
    While tcpClient.State <> 0
        Lblstatut.Caption = "statut : " & Str(tcpClient.State)
        DoEvents
    Wend
    recu = True
    CmdEnvoi.Enabled = False
    FormCrack.MousePointer = vbDefault
    
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

    tcpClient.PeekData strData
    If OptCertaine = True Then
        If Mid(strData, 10, 3) = "200" Then
            compteF = compteF + 1
            txtcompteF = Str(compteF)
            If total = True Then
                ListOK.AddItem (ligne)
            End If
        End If
    Else
        If Mid(strData, 10, 3) <> "404" Then
            compteF = compteF + 1
            txtcompteF = Str(compteF)
            If total = True Then
                ListOK.AddItem (ligne)
            End If
        End If
    End If
    If Len(TxtOutput) > 30000 Then
        TxtOutput.Text = Right(TxtOutput.Text, 28000)
    End If
    If total = False Then
        TxtOutput.Text = TxtOutput.Text & strData & vbCrLf
    End If
    TxtOutput.SelStart = Len(TxtOutput.Text)
    recu = True

End Sub

Private Sub TxtRemoteHost_GotFocus()
    TxtRemoteHost.SelStart = 0
    TxtRemoteHost.SelLength = Len(TxtRemoteHost.Text)
End Sub

Private Sub TxtRemotePort_GotFocus()
    TxtRemotePort.SelStart = 0
    TxtRemotePort.SelLength = Len(TxtRemotePort.Text)
End Sub
