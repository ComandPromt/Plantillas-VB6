VERSION 5.00
Begin VB.Form FormPing 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crack80"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtServeur 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox TimeOut 
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
      TabIndex        =   2
      Text            =   "500"
      ToolTipText     =   "Cette valeur permet de fixer la valeur de TimeOut  lors du contrôle d'éxistence de la cible"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdPing 
      BackColor       =   &H00808080&
      Caption         =   "Ping"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Height          =   615
      Index           =   0
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   2160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   2160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox Text1 
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
      Index           =   4
      Left            =   2160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4320
      Width           =   3975
   End
   Begin VB.TextBox Text1 
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
      Index           =   5
      Left            =   2160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox txtIP 
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
      Left            =   2160
      TabIndex        =   0
      Text            =   "195.154.24.1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Nom de Machine"
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
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Adresse IP"
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
      TabIndex        =   18
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "TimeOut (ms)"
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
      TabIndex        =   17
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Pointeur données"
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
      TabIndex        =   16
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Chaine envoyée"
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
      TabIndex        =   15
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Taille données"
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
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Délai"
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
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Adresse"
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
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Statut"
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
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "FormPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPing_Click()
Dim ECHO As ICMP_ECHO_REPLY
Dim pos As Integer
Dim retour As Long
Dim IP As String

timeout_ping = Val(TimeOut.Text)
For pos = 0 To 5
    Text1(pos).Text = ""
Next pos
FormPing.Refresh

WinsockInit

If txtIP <> "" Then
    txtServeur.Text = DottedIPToDNS(txtIP.Text)
Else
    IP = txtServeur.Text
    txtIP.Text = DNSToDottedIP(IP)
End If

Call Ping(txtIP.Text, ECHO)
Text1(0) = GetStatusCode(ECHO.status)
Text1(1) = ECHO.Address
Text1(2) = ECHO.RoundTripTime & " ms"
Text1(3) = ECHO.DataSize & " bytes"
If Left$(ECHO.Data, 1) <> Chr$(0) Then
pos = InStr(ECHO.Data, Chr$(0))
Text1(4) = Left$(ECHO.Data, pos - 1)
End If
Text1(5) = ECHO.DataPointer
WSACleanup
End Sub

Private Sub Form_Activate()
    Call couleur(Me)
End Sub
Private Sub CmdQuit_Click()
    FormPing.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub

Private Sub TimeOut_GotFocus()
    TimeOut.SelStart = 0
    TimeOut.SelLength = Len(TimeOut.Text)
End Sub

Private Sub txtIP_GotFocus()
    txtIP.SelStart = 0
    txtIP.SelLength = Len(txtIP.Text)
End Sub

Private Sub txtServeur_GotFocus()
    txtServeur.SelStart = 0
    txtServeur.SelLength = Len(txtServeur.Text)
End Sub
