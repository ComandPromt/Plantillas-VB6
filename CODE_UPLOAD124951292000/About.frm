VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Box"
   ClientHeight    =   1410
   ClientLeft      =   3570
   ClientTop       =   3210
   ClientWidth     =   2715
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2715
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"About.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Visible = False
End Sub
