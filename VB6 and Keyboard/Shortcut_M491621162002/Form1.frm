VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shortcut Maker"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "E"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Works when any part of form has focus..."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "To close this program press:"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Shortcut Key:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Enter a shortcut key"
Else
Timer1.Enabled = True
End If
Command1.Enabled = False
End Sub
Private Sub Command1_KeyPress(KeyAscii As Integer)
Timer1.Enabled = True
    If Timer1.Enabled = True Then
    If KeyAscii = Asc(Text1.Text) Then
    End
    Else
Timer1.Enabled = False
    End If
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Timer1.Enabled = True
    If Timer1.Enabled = True Then
    If KeyAscii = Asc(Text1.Text) Then
    End
    Else
Timer1.Enabled = False
    End If
    End If
End Sub
Private Sub Text1_GotFocus()
Command1.Enabled = True
Text1.BackColor = vbWhite
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
Timer1.Enabled = True
    If Timer1.Enabled = True Then
    If KeyAscii = Asc(Text1.Text) Then
    End
    Else
Timer1.Enabled = False
    End If
    End If
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = vbYellow
Command1_Click 'in case they forget to set :p
End Sub

