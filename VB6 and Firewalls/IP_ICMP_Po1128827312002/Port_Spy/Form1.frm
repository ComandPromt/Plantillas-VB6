VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Violation"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OnTop As New clsOnTop

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
OnTop.MakeTopMost hWnd
End Sub
