VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChecker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Checker"
   ClientHeight    =   2430
   ClientLeft      =   8250
   ClientTop       =   1530
   ClientWidth     =   3180
   Icon            =   "frmChecker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   3180
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3836
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheWordList

Function CheckWord(TheWord) As Boolean

A = TheWordList
B = InStr(LCase(A), LCase(TheWord))

If B = 0 Then

c = Right(TheWord, 1)
If LCase(c) = "s" Then
e = left(TheWord, Len(TheWord) - 1)
d = InStr(LCase(A), LCase(e))
If d = 0 Then CheckWord = False: Exit Function Else CheckWord = True: Exit Function

Else
CheckWord = False
Exit Function
End If

Else

CheckWord = True

End If

End Function

Function OpenFile(FileName)
On Error GoTo NahtMe
Open FileName For Input As #1
OpenFile = Input$(LOF(1), 1)
Close #1
Exit Function
NahtMe:
OpenFile = ""
Exit Function
End Function

Private Sub Form_Load()
TheWordList = OpenFile(App.Path & "\words.txt")
If Command1.Caption = "Check" Then
Command1.Caption = "Stop"
TreeView1.Nodes.Clear
TreeView1.Nodes.Add , , "x01", "Misspelled Words"
pos = 0
Do
A = frmMain.Text1.Text
B = Trim(Right(A, Len(A) - pos))
c = InStr(B, " ")
If c = 0 Then
d = CheckWord(B)
If d = "False" Then TreeView1.Nodes.Add "x01", 4, "y" & pos & "~" & Len(B), B: TreeView1.Nodes.Item("y" & pos & "~" & Len(B)).EnsureVisible
Caption = "Spell Checker - 100%"
Exit Do
Else
e = Trim(left(B, c - 1))
f = CheckWord(e)
If f = "False" Then TreeView1.Nodes.Add "x01", 4, "y" & pos & "~" & Len(e), e: TreeView1.Nodes.Item("y" & pos & "~" & Len(e)).EnsureVisible
pos = pos + Len(e) + 1
G = pos / Len(A)
H = Int(G * 100)
Caption = "Spell Checker - " & H & "%"
End If
If Command1.Caption = "Check" Then Exit Do
DoEvents
Loop
Command1.Caption = "Check"
Else
Caption = "Spell Checker"
Command1.Caption = "Check"
End If
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub TreeView1_Click()
On Error Resume Next
A = TreeView1.SelectedItem.Key
If Not left(A, 1) = "y" Then Exit Sub
B = Right(A, Len(A) - 1)
c = InStr(B, "~")
d = left(B, c - 1)
e = Right(B, Len(B) - c)
frmMain.Text1.SetFocus
frmMain.Text1.SelStart = d
frmMain.Text1.SelLength = e
End Sub
