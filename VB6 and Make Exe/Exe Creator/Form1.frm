VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Miniature EXE Creator"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Create"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   3975
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "c:\output.exe"
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Output file:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "What is it?"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
      Begin VB.Label Label4 
         Caption         =   $"Form1.frx":0000
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "EXE Information"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
      Begin VB.TextBox txtBody 
         Height          =   2175
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Body:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
Call MakeEXE
End Sub



Private Sub MakeEXE()
Dim tempInt As Integer
On Error GoTo ErrorHandler
Open txtOutput.Text For Binary Access Write As #1

'// Put some stuff in the file so the cpu knows what to do
'basically just some assembly machine code.//
Put #1, 1, 180
Put #1, 2, 9
Put #1, 3, 186
Put #1, 4, 9
Put #1, 5, 1
Put #1, 6, 205
Put #1, 7, 33
Put #1, 8, 195
Put #1, 9, 32

'// Insert your message //
For i = 1 To Len(txtBody.Text)
    tempInt = Asc(Mid(txtBody.Text, i, 1))
    Put #1, i + 9, tempInt
Next i

'// Put the footer //
Put #1, Len(txtBody.Text) + 10, 36
Close #1
MsgBox "EXE Compiled and Linked.", vbInformation, "Finished."

Exit Sub

ErrorHandler:
MsgBox "There was an error.", vbCritical, "Error"
End Sub

Private Sub Form_Load()
MsgBox "Please do not vote for this, just please leave a comment. This code was written by me, but the idea was from Vbmew (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=2232&lngWId=3), so please check it out!", vbInformation, "About"
End Sub
