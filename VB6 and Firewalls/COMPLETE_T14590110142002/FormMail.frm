VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormMail 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crack80"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDate 
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
      Left            =   3120
      TabIndex        =   4
      Text            =   "1 Jan 85 10:05:12"
      ToolTipText     =   "objet du message"
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox TxtObjet 
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
      Left            =   3120
      TabIndex        =   3
      Text            =   "Message"
      ToolTipText     =   "objet du message"
      Top             =   1560
      Width           =   4935
   End
   Begin MSWinsockLib.Winsock tcpMail 
      Left            =   5280
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TxtMessage 
      BackColor       =   &H00C0C0C0&
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Texte du message"
      Top             =   2640
      Width           =   7935
   End
   Begin VB.TextBox TxtMailDest 
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
      Left            =   3120
      TabIndex        =   2
      Text            =   "blaireau@dommage.fr"
      ToolTipText     =   "E-mail du blaireau"
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox TxtMailExp 
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
      Left            =   3120
      TabIndex        =   1
      Text            =   "bill@free.fr"
      ToolTipText     =   "Faux expéditeur..."
      Top             =   600
      Width           =   4935
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
      Left            =   3120
      TabIndex        =   0
      Text            =   "mail.oreka.fr"
      ToolTipText     =   "Machine servant de relais"
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Retour au menu"
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton CmdMail 
      BackColor       =   &H00808080&
      Caption         =   "Envoyer"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Envoie le mail"
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Date :"
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
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Objet :"
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
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label statut 
      BackColor       =   &H00404040&
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   7935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "E-Mail du destinataire :"
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
      TabIndex        =   10
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "E-Mail de l'expéditeur :"
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
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Serveur SMTP :"
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
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FormMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdMail_Click()
Dim pos As Integer
' Contrôle des zones envoyées
pos = InStr(TxtMailDest.Text, " ")
If pos > 0 Then
    statut.Caption = "Erreur mail destinataire"
    TxtMailDest.SetFocus
    Exit Sub
End If
pos = InStr(TxtMailExp.Text, " ")
If pos > 0 Then
    statut.Caption = "Erreur mail expéditeur"
    TxtMailExp.SetFocus
    Exit Sub
End If

tcpMail.RemotePort = 25
tcpMail.RemoteHost = TxtMachine.Text

    Me.MousePointer = vbHourglass
    statut.Caption = "Envoi du mail en cours !"
    
    If tcpMail.State <> 0 Then
        tcpMail.Close
    End If
        tcpMail.Connect
        While tcpMail.State <> 7
            If tcpMail.State = 9 Then
                statut.Caption = "Connexion impossible !"
                Me.MousePointer = vbNormal
                Exit Sub
            End If
            DoEvents
        Wend
        ' envoi du HELO
        statut.Caption = "HELO"
        tcpMail.SendData ("HELO " & TxtMachine.Text & Chr$(13) & Chr$(10))
        Do While Reponse <> "250"
            DoEvents
        Loop
        Reponse = ""
        ' envoi Mail expediteur
        statut.Caption = TxtMailExp.Text
        tcpMail.SendData ("MAIL FROM: " & TxtMailExp.Text & Chr$(13) & Chr$(10))
        Do While Reponse <> "250"
            DoEvents
        Loop
        Reponse = ""
        ' envoi Mail destinataire
        statut.Caption = TxtMailDest.Text
        tcpMail.SendData ("RCPT TO: " & TxtMailDest.Text & Chr$(13) & Chr$(10))
        Do While Reponse <> "250"
            DoEvents
        Loop
        Reponse = ""
        ' envoi Data
        statut.Caption = "DATA"
        tcpMail.SendData ("DATA" & Chr$(13) & Chr$(10))
        Do While Reponse <> "354"
            DoEvents
        Loop
        Reponse = ""
        ' envoi du message
        statut.Caption = "Message en cours d'envoi !"
        tcpMail.SendData ("SUBJECT: " & TxtObjet.Text & Chr$(13) & Chr$(10) & _
        "DATE: " & TxtDate.Text & _
        "FROM: " & TxtMailExp.Text & " <" & TxtMailExp.Text & ">" & Chr$(13) & Chr$(10) & _
        "TO: " & TxtMailDest.Text & " <" & TxtMailDest.Text & ">" & Chr$(13) & Chr$(10) & _
        TxtMessage.Text & Chr$(13) & Chr$(10))
        
        'Winsock1.SendData "FROM: " & From & " <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10)
        'Winsock1.SendData "TO: " & MAIL_TO & " <" & RCPT_TO & ">" & Chr$(13) & Chr$(10)
        
        ' envoi de la fin de message
        tcpMail.SendData (Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10))
        Do While Reponse <> "250"
            DoEvents
        Loop
        Reponse = ""
        tcpMail.SendData ("QUIT" & Chr$(13) & Chr$(10))
        Do While Reponse <> "221"
            DoEvents
        Loop
        
        tcpMail.Close
        While tcpMail.State <> 0
            If tcpMail.State = 9 Then
                statut.Caption = "Déconnexion impossible !"
                Me.MousePointer = vbNormal
                Exit Sub
            End If
            DoEvents
        Wend
  
    statut.Caption = "Envoi du mail terminé!"
    Me.MousePointer = vbNormal
End Sub
Private Sub CmdQuit_Click()
    If tcpMail.State <> 0 Then
        tcpMail.Close
    End If
    FormMail.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub
Private Sub Form_Activate()
Dim ddate
Dim dday
Dim mmonth
Dim yyear
Dim ttime
Dim mois
mois = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    
    ttime = Time
    ddate = Date
    dday = Day(ddate)
    mmonth = Month(ddate)
    yyear = Year(ddate)
    TxtDate.Text = dday & " " & mois(Val(mmonth) - 1) & " " & yyear & " " & ttime
    Call couleur(Me)
End Sub
Private Sub TxtDate_GotFocus()
    TxtDate.SelStart = 0
    TxtDate.SelLength = Len(TxtDate.Text)
End Sub
Private Sub TxtMachine_GotFocus()
    TxtMachine.SelStart = 0
    TxtMachine.SelLength = Len(TxtMachine.Text)
End Sub
Private Sub TxtNomExp_GotFocus()
    TxtNomExp.SelStart = 0
    TxtNomExp.SelLength = Len(TxtNomExp.Text)
End Sub
Private Sub TxtMailExp_GotFocus()
    TxtMailExp.SelStart = 0
    TxtMailExp.SelLength = Len(TxtMailExp.Text)
End Sub
Private Sub TxtNomDest_GotFocus()
    TxtNomDest.SelStart = 0
    TxtNomDest.SelLength = Len(TxtNomDest.Text)
End Sub
Private Sub TxtMailDest_GotFocus()
    TxtMailDest.SelStart = 0
    TxtMailDest.SelLength = Len(TxtMailDest.Text)
End Sub
Private Sub TxtObjet_GotFocus()
    TxtObjet.SelStart = 0
    TxtObjet.SelLength = Len(TxtObjet.Text)
End Sub
Private Sub tcpMail_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

    On Error GoTo ErrTcp
    
    tcpMail.GetData strData
    Reponse = Mid(strData, 1, 3)
    Exit Sub
    
ErrTcp:
    statut.Caption = "Erreur !!!!!"
    tcpMail.Close
    Me.MousePointer = vbDefault
End Sub
