VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "                           Josh's NotePad Example"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open a text file or html..."
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "0"
      Top             =   3120
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   3720
      Y2              =   3960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Max Character Limit:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
warning = MsgBox("Your current data is now going to be erased! Make sure that you have already saved it.", vbOKCancel, "Warning!")
'warning
If warning = vbCancel Then
Else
If warning = vbOK Then
RichTextBox1.Text = ""
End If
End If


End Sub

Private Sub Command2_Click()
CommonDialog1.Filter = "Text files|*.txt|HTML Files (*.html)|*.html|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    RichTextBox1.LoadFile CommonDialog1.filename, rtfText
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowSave
    RichTextBox1.SaveFile CommonDialog1.filename, rtfText

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
RichTextBox1.MaxLength = 0
End Sub

Private Sub Text1_Change()
On Error Resume Next
RichTextBox1.MaxLength = Text1.Text
End Sub
