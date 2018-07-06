VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2505
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   2505
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Height          =   375
      Left            =   1455
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdFind_Click()
Dim textfound As Integer

   
cmdFindNext.Enabled = True

   
frmMain.Text1.find (Text1.Text)
frmMain.Text1.SetFocus

   
textfound = frmMain.Text1.find(Text1.Text)
If textfound = -1 Then
MsgBox vbCr & "Text could not be found.", vbInformation, "SDI Word"
End If
End Sub

Private Sub cmdFindNext_Click()
frmMain.Text1.SetFocus

   
frmMain.Text1.find (Text1.Text), frmMain.Text1.SelStart + 1
End Sub

Private Sub Form_Load()

End Sub
