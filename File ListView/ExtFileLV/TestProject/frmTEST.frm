VERSION 5.00
Object = "*\A..\ExtFileLV.vbp"
Begin VB.Form frmTEST 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin ExtFileLV.FileLV lvLIST 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      Path            =   "C:"
   End
End
Attribute VB_Name = "frmTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Resize()
    On Error Resume Next
    With Me
        .lvLIST.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
End Sub
Private Sub lvLIST_DblClick()
    On Error Resume Next
    If lvLIST.ItemType = itFOLDER Then
        lvLIST.Path = lvLIST.Path & "\" & lvLIST.SelectedItem.Text
    End If
End Sub
