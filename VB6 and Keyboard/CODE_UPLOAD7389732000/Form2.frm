VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Report as of January 320 B.C."
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Press SHIFT + M to go back to main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I submit this code in a .ZIP file because I do not like text
'I mean I do not like Copy and Paste.


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyM And Shift = 1 Then Form1.Show: Unload Me
End Sub


Private Sub Form_Load()
KeyPreview = True
End Sub
