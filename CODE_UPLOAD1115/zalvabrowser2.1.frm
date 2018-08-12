VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Password check"
   ClientHeight    =   585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2580
   Icon            =   "zalvabrowser2.1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   2580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then

If Text1.Text = "browser" Then
Unload Form1
Form2.Show
Unload Me
Unload Form1
End If

End If
End Sub
