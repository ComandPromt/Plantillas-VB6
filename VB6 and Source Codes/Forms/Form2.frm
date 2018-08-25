VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9255
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   9255
   Begin VB.TextBox Text1 
      Height          =   7335
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   30
      Width           =   9195
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Text1.Width = Form2.Width - 200
Text1.Height = Form2.Height - 500
End Sub
