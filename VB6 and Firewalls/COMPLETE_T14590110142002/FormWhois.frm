VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormWhois 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crack80"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbnic 
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
      Height          =   345
      ItemData        =   "FormWhois.frx":0000
      Left            =   120
      List            =   "FormWhois.frx":001C
      TabIndex        =   5
      Text            =   "rs.internic.net"
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox TxtWhois 
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
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "Descriptif du serveur demandé"
      Top             =   1080
      Width           =   6615
   End
   Begin MSWinsockLib.Winsock tcpWhois 
      Left            =   2040
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Retour au menu"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton CmdWhois 
      BackColor       =   &H00808080&
      Caption         =   "Whois"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Whois... qui est qui ?"
      Top             =   600
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
      Left            =   2160
      TabIndex        =   1
      Text            =   "oreka.fr"
      ToolTipText     =   "Machine sur laquelle on souhaite des informations "
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Nom de machine :"
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
      Width           =   1935
   End
End
Attribute VB_Name = "FormWhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdQuit_Click()
    FormWhois.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub

Private Sub CmdWhois_Click()
    Me.MousePointer = vbHourglass
    If tcpWhois.State <> 0 Then
        tcpWhois.Close
    End If
    tcpWhois.RemotePort = 43
    If Cmbnic.Text = "" Then
        Cmbnic.Text = "rs.internic.net"
    End If
    tcpWhois.RemoteHost = Cmbnic.Text
    TxtWhois.Text = ""
    tcpWhois.Connect
    While tcpWhois.State <> 7
        If tcpWhois.State = 9 Then
            TxtWhois = TxtWhois & " " & Cmbnic.Text & " non disponible !" & vbCrLf
            Me.MousePointer = vbNormal
            Exit Sub
        End If
        DoEvents
    Wend
    tcpWhois.SendData TxtMachine & vbCrLf
    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Activate()
    Call couleur(Me)
End Sub

Private Sub tcpWhois_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

    On Error GoTo ErrTcp
    
    tcpWhois.GetData strData
    strData = Replace(strData, Chr$(10), vbCrLf)
    TxtWhois = TxtWhois & strData
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrTcp:
    TxtWhois = TxtWhois & "Erreur !!!!!" & vbCrLf
    tcpWhois.Close
    Me.MousePointer = vbDefault
End Sub

Private Sub TxtMachine_GotFocus()
    TxtMachine.SelStart = 0
    TxtMachine.SelLength = Len(TxtMachine.Text)
End Sub

Private Sub tcpWhois_Close()
    TxtWhois.SelStart = Len(TxtWhois.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If tcpWhois.State <> 0 Then
    tcpWhois.Close
End If
End Sub
