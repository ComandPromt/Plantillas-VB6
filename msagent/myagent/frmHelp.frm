VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   7995
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmHelp.frx":0000
      Top             =   60
      Width           =   9135
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strF As String

    'load a file into a string
    
    Open "C:\auotexec.bat" For Binary As #1
    strF = Space$(LOF(1))
    Get #1, , strF
    Close #1
    'display file
    
    Text1.Text = strF


End Sub
