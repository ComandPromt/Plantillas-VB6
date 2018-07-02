VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormIrc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crack80"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox checkTrace 
      BackColor       =   &H00000000&
      Caption         =   "Trace"
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
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton CmdCmd 
      BackColor       =   &H00808080&
      Caption         =   "envoi commande"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "envoi une commande au serveur"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtCommande 
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
      Height          =   405
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "ex : join #toto     who #toto     whois nick     names #toto     list"
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox TxtYourNick 
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
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Text            =   "crack80"
      ToolTipText     =   "Serveur de Chat sur lequel on veut se connecter"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox TxtIrc 
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
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      ToolTipText     =   "Descriptif du serveur demandé"
      Top             =   1560
      Width           =   9015
   End
   Begin MSWinsockLib.Winsock tcpIrc 
      Left            =   3840
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7680
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Retour au menu"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton CmdIrc 
      BackColor       =   &H00808080&
      Caption         =   "Connexion"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4560
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "connexion au serveur IRC"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox TxtMachine 
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
      Height          =   405
      Left            =   2280
      TabIndex        =   1
      Text            =   "wanadoo.entrechat.net"
      ToolTipText     =   "Serveur de Chat sur lequel on veut se connecter"
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label statut2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label statut 
      BackColor       =   &H00000000&
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   5535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Commande :"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Votre Nick :"
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
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Serveur de Chat :"
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
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FormIrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub checkTrace_Validate(Cancel As Boolean)
    Call ecrire_log("Activation Trace : IRC")
    Call ecrire_log(Now)
End Sub

Private Sub CmdCmd_Click()
    If checkTrace.Value = 1 Then
        Call ecrire_log("==========> " & txtCommande.Text)
    End If
    statut.Caption = "Envoi de " & txtCommande.Text
    tcpIrc.SendData txtCommande.Text & " " & vbCrLf
    statut.Caption = "Commande exécutée !"
    txtCommande.SetFocus
End Sub
Private Sub Form_Activate()
    Call couleur(Me)
End Sub

Private Sub CmdQuit_Click()
    ' envoie du QUIT si on est connecte
    If tcpIrc.State = 7 Then
        txtCommande.Text = "QUIT"
        Call CmdCmd_Click
    End If
    Call ecrire_log("Fin IRC : " & Now)
    FormIrc.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub

Private Sub CmdIrc_Click()
Dim dpong, fpong As Integer
Dim strpong As String
    
    Me.MousePointer = vbHourglass
    If tcpIrc.State <> 0 Then
        tcpIrc.Close
    End If
    tcpIrc.RemotePort = 6667
    
    tcpIrc.RemoteHost = TxtMachine.Text
    statut.Caption = "Connexion en cours..."
    tcpIrc.Connect
    While tcpIrc.State <> 7
        If tcpIrc.State = 9 Then
            TxtIrc = TxtMachine.Text & " non disponible !" & vbCrLf
            If checkTrace.Value = 1 Then
                Call ecrire_log(TxtIrc)
            End If
            Me.MousePointer = vbNormal
            Exit Sub
        End If
        DoEvents
    Wend
    statut.Caption = "Serveur connecté !"
    statut.Caption = "Envoi du nickname"
    If checkTrace.Value = 1 Then
        Call ecrire_log(TxtYourNick.Text & " connecté !")
    End If
    tcpIrc.SendData "NICK " & TxtYourNick.Text & vbCrLf
    statut2.Caption = Reponse
    Do While Mid(UCase(Reponse), 1, 4) <> "PING"
            statut.Caption = "Attente du PING..."
            statut2.Caption = statut2.Caption & "."
            If Len(statut2.Caption) > 35 Then
                statut2.Caption = Reponse
            End If
            statut2.Refresh
            DoEvents
    Loop
    strpong = Mid(Reponse, 7)
    statut.Caption = "Envoi de PONG :" & strpong
    tcpIrc.SendData "PONG :" & strpong & vbCrLf
    Reponse = ""
    
    'statut.Caption = "Envoi du user"
    tcpIrc.SendData "USER " & "test test test :crack80" & vbCrLf
    
    statut.Caption = "Séquence terminée"
    statut2.Caption = ""
    
    txtCommande.SetFocus
    Me.MousePointer = vbNormal
End Sub


Private Sub tcpIrc_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

    On Error GoTo ErrTcp
    
    tcpIrc.GetData strData
    strData = Replace(strData, Chr$(10), vbCrLf)
    TxtIrc = TxtIrc & strData & vbCrLf
    If Len(TxtIrc) > 32000 Then
        TxtIrc = Mid(TxtIrc, 2000, Len(TxtIrc) - 2000)
    End If
    TxtIrc.SelStart = Len(TxtIrc.Text)
    Reponse = strData
    ' ecriture du log si trace activee
    If checkTrace.Value = 1 Then
        Call ecrire_log(strData)
    End If
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrTcp:
    TxtIrc = TxtIrc & "Erreur !!!!!" & vbCrLf
    tcpIrc.Close
    Me.MousePointer = vbDefault
End Sub

Private Sub txtCommande_GotFocus()
    txtCommande.SelStart = 0
    txtCommande.SelLength = Len(txtCommande)
End Sub

Private Sub TxtMachine_GotFocus()
    TxtMachine.SelStart = 0
    TxtMachine.SelLength = Len(TxtMachine.Text)
End Sub

Private Sub tcpIrc_Close()
    TxtIrc.SelStart = Len(TxtIrc.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If tcpIrc.State <> 0 Then
    tcpIrc.Close
End If
End Sub
