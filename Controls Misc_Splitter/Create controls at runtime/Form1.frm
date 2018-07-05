VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Text            =   "Original"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Menu dg 
      Caption         =   "cg"
      Begin VB.Menu dgt 
         Caption         =   "dfg"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
'if you want to create control at runtime you must set
'index to 0
'like:  put command1 on form and set index to 0
If Index = 0 Then
Load Command1(Command1.Count + 1)
   Command1(Command1.Count).Left = 1400 + Command1.Count * 30
   Command1(Command1.Count).Top = 1000 * 1 + Command1.Count * 30
   Command1(Command1.Count).Caption = "NUMBER " & Command1.Count
   Command1(Command1.Count).Visible = True
   
Load Text1(Text1.Count + 1)
   Text1(Text1.Count).Left = Text1(Text1.Count).Width + 900 + Command1.Count * 30
   Text1(Text1.Count).Top = Text1(Text1.Count).Top + Command1.Count * 30
   Text1(Text1.Count).Text = "New textbox"
   Text1(Text1.Count).Visible = True
For I = 1 To 5
Load dgt(dgt.Count + 1)
   dgt(dgt.Count).Caption = "COPY" & I
   dgt(dgt.Count).Visible = True
   
 Next I
 
Dim dgt2b As Menu
Form1.PopupMenu dgt2b
dgt2b.Visible = True

ElseIf Index = 1 Then
MsgBox "You have created copy of command1"
End If
End Sub

Private Sub dgt_Click(Index As Integer)
 MsgBox dgt.Item(Index).Caption
End Sub

Private Sub Form_Load()

        Frame1.Width = Screen.Width + 100
        Frame1.Move -50, 0
       
End Sub
