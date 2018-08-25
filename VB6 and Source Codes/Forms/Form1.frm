VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   9330
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   120
      TabIndex        =   27
      Top             =   360
      Width           =   9135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Z"
      Height          =   285
      Index           =   25
      Left            =   7620
      TabIndex        =   26
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Y"
      Height          =   285
      Index           =   24
      Left            =   7320
      TabIndex        =   25
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   285
      Index           =   23
      Left            =   7020
      TabIndex        =   24
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "W"
      Height          =   285
      Index           =   22
      Left            =   6720
      TabIndex        =   23
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "V"
      Height          =   285
      Index           =   21
      Left            =   6420
      TabIndex        =   22
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "U"
      Height          =   285
      Index           =   20
      Left            =   6120
      TabIndex        =   21
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "T"
      Height          =   285
      Index           =   19
      Left            =   5820
      TabIndex        =   20
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S"
      Height          =   285
      Index           =   18
      Left            =   5520
      TabIndex        =   19
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R"
      Height          =   285
      Index           =   17
      Left            =   5220
      TabIndex        =   18
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Q"
      Height          =   285
      Index           =   16
      Left            =   4920
      TabIndex        =   17
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "P"
      Height          =   285
      Index           =   15
      Left            =   4620
      TabIndex        =   16
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O"
      Height          =   285
      Index           =   14
      Left            =   4320
      TabIndex        =   15
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N"
      Height          =   285
      Index           =   13
      Left            =   4020
      TabIndex        =   14
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "M"
      Height          =   285
      Index           =   12
      Left            =   3720
      TabIndex        =   13
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "L"
      Height          =   285
      Index           =   11
      Left            =   3420
      TabIndex        =   12
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "K"
      Height          =   285
      Index           =   10
      Left            =   3120
      TabIndex        =   11
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "J"
      Height          =   285
      Index           =   9
      Left            =   2820
      TabIndex        =   10
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I"
      Height          =   285
      Index           =   8
      Left            =   2520
      TabIndex        =   9
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "H"
      Height          =   285
      Index           =   7
      Left            =   2220
      TabIndex        =   8
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "G"
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   7
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F"
      Height          =   285
      Index           =   5
      Left            =   1620
      TabIndex        =   6
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E"
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   5
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "D"
      Height          =   285
      Index           =   3
      Left            =   1020
      TabIndex        =   4
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B"
      Height          =   285
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   285
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   180
      TabIndex        =   0
      Top             =   3930
      Width           =   7785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
List1.Clear

For i = 0 To File1.ListCount - 1
    If Mid(File1.List(i), 1, 1) = Command1(Index).Caption Then
        List1.AddItem File1.List(i)
    End If
Next i

End Sub

Private Sub Form_Load()
File1.Path = "C:\WINDOWS\Desktop\VisualBasic\Source Codes"
End Sub

Private Sub List1_Click()

    crlf$ = Chr(13) & Chr(10)
    Form2.Text1.Text = ""
    ThePath = File1.Path & "\" & List1.Text
    Open ThePath For Input As #1
    While Not EOF(1)
        Line Input #1, file_data$
        Form2.Text1.Text = Form2.Text1.Text & file_data$ & crlf$
    Wend
    Close #1


Form2.Show
End Sub
