VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Listbox tricks"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Check Duplicates"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   4575
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Shuffle Items"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort Items"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Items"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
List1.Visible = False
Dim p As Integer
p = List1.ListCount - 1
For i = 0 To List1.ListCount - 1
    If i > p Then
        List1.ListIndex = 0
        List1.Visible = True
        Exit Sub
        End If
    List1.ListIndex = i
    If List1.Text = "hello2" Then
        List1.RemoveItem List1.ListIndex
        i = i - 1
        p = p - 1
        End If
Next i
List1.ListIndex = 0
List1.Visible = True
End Sub

Private Sub Command2_Click()
Dim q As Integer
Dim store As String
List2.Visible = False
For i = 0 To List2.ListCount - 1
    q = i + 1
    List2.ListIndex = i
    Text1.Text = Mid(List2.Text, 1, 1)
    If q > List2.ListCount - 1 Then
        List2.ListIndex = 0
        List2.Visible = True
        Exit Sub
        End If
    List2.ListIndex = q
    Text2.Text = Mid(List2.Text, 1, 1)
    If Text2.Text = Text1.Text + 1 Then
    Else
        Text2.Text = Text1.Text + 1
        List2.ListIndex = q
        store = List2.Text
        List2.RemoveItem List2.ListIndex
        List2.AddItem Text2.Text & Mid(store, 2, Len(store)), q
    End If
Next i
List2.ListIndex = 0
List2.Visible = True
End Sub

Private Sub Command3_Click()
Dim r As Integer
Dim storeit As String
r = List3.ListCount - 1
List3.Visible = False
For i = 0 To List3.ListCount - 1
    List3.ListIndex = i
    storeit = List3.Text
    List3.RemoveItem List3.ListIndex
    List3.AddItem storeit, Int(Rnd * r)
Next i
List3.ListIndex = 0
List3.Visible = True
End Sub

Private Sub Command4_Click()
Dim s As Integer
Dim storethis As String
s = List4.ListCount - 1
List4.Visible = False
For i = 0 To List4.ListCount - 1
    If i > s Then
        List4.ListIndex = 0
        List4.Visible = True
        Exit Sub
        End If
    List4.ListIndex = i
    storethis = List4.Text
    For m = 0 To List4.ListCount - 1
        If m > s Then
            m = 1
            GoTo thenextone
            End If
        If m = i Then
            GoTo cont
        Else
            List4.ListIndex = m
        End If
        If storethis = List4.Text Then
            List4.RemoveItem List4.ListIndex
            s = s - 1
        Else
        End If
cont:
    Next m
thenextone:
    storethis = ""
Next i
List4.ListIndex = 0
List4.Visible = True
End Sub

Private Sub Form_Load()
'Add stuff into listbox1
List1.AddItem "1"
List1.AddItem "hello2"
List1.AddItem "2"
List1.AddItem "hello2"
List1.AddItem "3"
List1.AddItem "hello2"
List1.AddItem "4"
List1.AddItem "hello2"
List1.AddItem "5"
List1.AddItem "hello2"
List1.AddItem "6"
List1.AddItem "hello2"
List1.AddItem "7"
List1.AddItem "hello2"
List1.AddItem "8"
List1.AddItem "hello2"
List1.AddItem "9"
List1.AddItem "hello2"
'Add stuff into listbox2
List2.AddItem "1 Hello1"
List2.AddItem "2 Hello2"
List2.AddItem "4 Hello3"
List2.AddItem "5 Hello4"
List2.AddItem "6 Hello5"
List2.AddItem "8 Hello6"
List2.AddItem "9 Hello7"
'Add stuff into listbox3
List3.AddItem "Place 1"
List3.AddItem "Place 2"
List3.AddItem "Place 3"
List3.AddItem "Place 4"
List3.AddItem "Place 5"
List3.AddItem "Place 6"
List3.AddItem "Place 7"
List3.AddItem "Place 8"
List3.AddItem "Place 9"
'Add stuff into listbox4
List4.AddItem "Dup 1"
List4.AddItem "Dup 1"
List4.AddItem "Dup 2"
List4.AddItem "Dup 2"
List4.AddItem "Dup 3"
List4.AddItem "Dup 3"
List4.AddItem "Dup 4"
List4.AddItem "Dup 4"
List4.AddItem "Dup 5"
List4.AddItem "Dup 5"
End Sub

