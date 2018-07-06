VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   2955
   ClientLeft      =   3840
   ClientTop       =   4140
   ClientWidth     =   4785
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Ending 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      ToolTipText     =   "What to Insert at End"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Begin 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "What to Insert at Begining"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Start at cursor location"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "DONE"
      Height          =   345
      Left            =   3240
      TabIndex        =   8
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Replace All To End"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Input what you wish to replace"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Whole Word"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox FindStuff 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Input what you wish to find"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Insert at end of line"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Insert at begining of line"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Find"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Replace"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Default As String
Dim Chck1, Chck2 As Integer
Dim sePosition As Long
Dim SameSpot As Boolean

Private Sub Begin_Change()
    If Len(Begin.Text) = 0 Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 0 Then
        sePosition = 0
    End If
End Sub

Private Sub cmdOK_Click()
    Form1.Hide
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim FindFlags As Long
    Screen.MousePointer = vbHourglass
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    If sePosition = Len(fMainForm.Text1.Text) - 1 Then GoTo restart
    If Default = FindStuff.Text And Check1.Value = Chck1 And Check2.Value = Chck2 Then
        If Check4.Value = 1 And SameSpot = False Then
            SameSpot = True
            sePosition = fMainForm.Text1.Find(FindStuff.Text, fMainForm.Text1.SelStart, , FindFlags)
        Else
            sePosition = fMainForm.Text1.Find(FindStuff.Text, sePosition + 1, , FindFlags)
        End If
        If sePosition >= 0 Then
            'fMainForm.SetFocus
            Label3.Caption = "Found"
            Command2.Enabled = True
        Else
            'Command2.Enabled = False
            Label3.Caption = "No More Found"
            Beep
            Command2.Enabled = False
            'fMainForm.SetFocus
        End If
    Else
restart:
        If Check4.Value = 1 Then
            sePosition = fMainForm.Text1.SelStart
        Else
            sePosition = -1
        End If
        sePosition = fMainForm.Text1.Find(FindStuff.Text, sePosition + 1, , FindFlags)
        If sePosition >= 0 Then
            'fMainForm.SetFocus
            Label3.Caption = "Found"
            Command2.Enabled = True
        Else
            Label3.Caption = "None Found!"
            Beep
            Command2.Enabled = False
        End If
    End If
    Screen.MousePointer = vbDefault
    Default = FindStuff.Text
    Chck1 = Check1.Value
    Chck2 = Check2.Value
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim FindFlags As Long
    FindFlags = Check1.Value * 4 + Check2.Value * 2
'   If sePosition >= Len(fMainForm.Text1.Text) - 2 Then GoTo restart
    If SameSpot = False Then sePosition = fMainForm.Text1.SelStart
    If Check3.Value = 1 Then
        While sePosition >= 0
            'If fMainForm.Text1.SelText = FindStuff.Text Then
            fMainForm.Text1.SelText = Text2.Text
            'End If
            sePosition = fMainForm.Text1.Find(FindStuff.Text, sePosition + Len(Text2.Text), , FindFlags)
            'fMainForm.Text1.SelText = Text2.Text
            'If sePosition >= Len(fMainForm.Text1.Text) - 1 Then GoTo leave
        Wend
        Beep
leave:
        Label3.Caption = "Done, Replaced All"
        Command2.Enabled = False
    Else
        If sePosition >= 0 Then
            'fMainForm.SetFocus
            fMainForm.Text1.SelText = Text2.Text
            sePosition = fMainForm.Text1.Find(FindStuff.Text, sePosition + Len(Text2.Text), , FindFlags)
            Beep
            'Label3.Caption = "Another Found, replace it?"
        Else
            Beep
            Label3.Caption = "Done, No more found!"
            Command2.Enabled = False
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
    Dim BeginPosition
    BeginPosition = 0
    Dim FindFlags As Long
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    fMainForm.Text1.SelStart = 0
        While BeginPosition < Len(fMainForm.Text1.Text)
            If BeginPosition = 0 Then
                fMainForm.Text1.SelText = Begin.Text
            Else
                fMainForm.Text1.SelText = vbNewLine & Begin.Text
            End If
            BeginPosition = fMainForm.Text1.Find(Chr(10), BeginPosition + Len(Begin.Text + vbNewLine), , FindFlags)
        Wend
End Sub

Private Sub Command4_Click()
    Dim BeginPosition
    BeginPosition = 0
    Dim FindFlags As Long
    fMainForm.Text1.SelStart = 0
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    BeginPosition = fMainForm.Text1.Find(Chr(10), BeginPosition + Len(Text2.Text), , FindFlags)
    While BeginPosition <= Len(fMainForm.Text1.Text)
        If BeginPosition = 0 Then
            fMainForm.Text1.SelText = Ending.Text
        Else
            fMainForm.Text1.SelText = Ending.Text & vbNewLine
        End If
        BeginPosition = fMainForm.Text1.Find(Chr(10), BeginPosition + Len(Ending.Text + vbNewLine), , FindFlags)
        If fMainForm.Text1.SelStart = Len(fMainForm.Text1.Text) Then
            fMainForm.Text1.SelText = Ending.Text
            Exit Sub
        End If
    Wend
End Sub

Private Sub Ending_Change()
    If Len(Ending.Text) = 0 Then
        Command4.Enabled = False
    Else
        Command4.Enabled = True
    End If
End Sub

Private Sub FindStuff_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Deactivate()
    SameSpot = False
End Sub

Private Sub Form_GotFocus()
    SameSpot = False
End Sub

Private Sub Form_Load()
    SameSpot = False
    Label3.Caption = "Status"
    Default = ""
    Chck1 = 0
    Chck2 = 0
    'FindStuff.Text = fMainForm.Text1.SelText
    'FindStuff.SetFocus
End Sub

Private Sub Form_LostFocus()
    SameSpot = False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Command2.Enabled = True Then Call Command2_Click
    End If
End Sub
