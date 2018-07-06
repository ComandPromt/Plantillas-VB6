VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find and Replace"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdSelection 
         Caption         =   "."
         Height          =   315
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         ToolTipText     =   "Get selected text"
         Top             =   480
         Width           =   135
      End
      Begin VB.CommandButton cmdSelection 
         Caption         =   "."
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   11
         ToolTipText     =   "Get selected text"
         Top             =   1080
         Width           =   135
      End
      Begin VB.CheckBox ChWord 
         Caption         =   "Whole word"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox ChCase 
         Caption         =   "Match Case"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cboReplace 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cboFind 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Replace..."
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Find..."
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Height          =   330
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   330
      Left            =   3120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   330
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'pretty standard RTF case sensitive search and replace
'except that the replace function is linked to the undo class
Dim st As Long
Dim mmatchCase As Integer
Dim mWholeword As Integer
Dim found As Long
Dim vStrPos As Long
Dim StopNow As Boolean
Dim Searching As Boolean
Private Sub cboFind_Change()
    If Len(Trim(cboFind.Text)) = 0 Then
        cmdFindNext.Enabled = False
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    Else
        cmdFindNext.Enabled = True
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
    End If
    st = 0
End Sub
Private Sub cmdCancel_Click()
    If Searching Then
        StopNow = True
        Searching = False
    Else
        Unload Me
        frmMain.EditEnable
    End If
End Sub
Private Sub cmdFindNext_Click()
    If ChWord.Value = 1 Then
        mWholeword = 2
    Else
        mWholeword = 0
    End If
    If ChCase.Value = 1 Then
        mmatchCase = 4
    Else
        mmatchCase = 0
    End If
    'No dupes thanks
    vStrPos = SendMessageByString&(cboFind.hwnd, CB_FINDSTRINGEXACT, 0, cboFind.Text)
    If vStrPos - 1 Then
        cboFind.AddItem cboFind.Text
    End If
    With frmMain.RTF
        LockWindowUpdate frmMain.hwnd
        found = .Find(cboFind.Text, st, , mWholeword Or mmatchCase)
        If found <> -1 Then
            st = found + Len(cboFind.Text)
            cmdReplace.Enabled = True
            .SetFocus
        Else
            st = 0
            cmdReplace.Enabled = False
            MsgBox "Text not found."
        End If
        LockWindowUpdate 0
    End With
End Sub
Private Sub cmdReplace_Click()
    If frmMain.RTF.SelText = cboFind.Text Then
        frmMain.Undo.InsertText cboReplace.Text
    Else
        cmdFindNext_Click
        If frmMain.RTF.SelText = cboFind.Text Then frmMain.Undo.InsertText cboReplace.Text
    End If
    'No dupes thanks
    vStrPos = SendMessageByString&(cboFind.hwnd, CB_FINDSTRINGEXACT, 0, cboFind.Text)
    If vStrPos = -1 Then
        cboFind.AddItem cboFind.Text
    End If
    vStrPos = SendMessageByString&(cboReplace.hwnd, CB_FINDSTRINGEXACT, 0, cboReplace.Text)
    If vStrPos = -1 Then
        cboReplace.AddItem cboReplace.Text
    End If
    FileChanged = True
End Sub
Private Sub cmdReplaceAll_Click()
    Dim count As Long, beginST As Long
    Searching = True
    Screen.MousePointer = 11
    If ChWord.Value = 1 Then
        mWholeword = 2
    Else
        mWholeword = 0
    End If
    If ChCase.Value = 1 Then
        mmatchCase = 4
    Else
        mmatchCase = 0
    End If
    vStrPos = SendMessageByString&(cboFind.hwnd, CB_FINDSTRINGEXACT, 0, cboFind.Text)
    If vStrPos = -1 Then
        cboFind.AddItem cboFind.Text
    End If
    vStrPos = SendMessageByString&(cboReplace.hwnd, CB_FINDSTRINGEXACT, 0, cboReplace.Text)
    If vStrPos = -1 Then
        cboReplace.AddItem cboReplace.Text
    End If
    NoStatusUpdate = True
    With frmMain.RTF
        .SelStart = 0
        beginST = .SelStart
        LockWindowUpdate frmMain.hwnd
        Do
            DoEvents
            If StopNow Then Exit Do
            found = .Find(cboFind.Text, st, , mWholeword Or mmatchCase)
            If found <> -1 Then
                st = found + Len(cboFind.Text)
                count = count + 1
                frmMain.Undo.InsertText cboReplace.Text, False
            Else
                st = 0
                cmdReplace.Enabled = False
                If count = 0 Then
                    MsgBox "Text not found."
                    Exit Do
                Else
                    FileChanged = True
                    Exit Do
                End If
            End If
            If StopNow Then Exit Do
        Loop
        StopNow = False
        .SelStart = beginST
        .SelLength = 0
        Searching = False
        NoStatusUpdate = False
        If count > 0 Then frmMain.Undo.UpdateStateChange
        Screen.MousePointer = 0
        MsgBox count & IIf(count = 1, " item replaced", " items replaced")
        LockWindowUpdate 0
    End With
End Sub
Private Sub Command1_Click()
    cboFind.Text = frmMain.RTF.SelText
End Sub
Private Sub Command2_Click()
    cboReplace.Text = frmMain.RTF.SelText
End Sub
Private Sub cmdSelection_Click(Index As Integer)
    'get selection from frmMain.RTF
    If Index = 0 Then
        cboFind.Text = frmMain.RTF.SelText
    Else
        cboReplace.Text = frmMain.RTF.SelText
    End If
End Sub
Private Sub Form_Load()
    Dim v As Integer, temp As String
    cboFind.BackColor = frmMain.RTF.BackColor
    cboReplace.BackColor = frmMain.RTF.BackColor
    'load up old searches/replaces from registry
    For v = 0 To 9
        temp = GetSetting("PSST SOFTWARE\" + App.Title, "FindText", "Text" + Str(v))
        If temp <> "" Then
            vStrPos = SendMessageByString&(cboFind.hwnd, CB_FINDSTRINGEXACT, 0, temp)
            If vStrPos = -1 Then
                cboFind.AddItem temp
            End If
        End If
        temp = GetSetting("PSST SOFTWARE\" + App.Title, "ReplaceText", "Text" + Str(v))
        If temp <> "" Then
            vStrPos = SendMessageByString&(cboReplace.hwnd, CB_FINDSTRINGEXACT, 0, temp)
            If vStrPos = -1 Then
                cboReplace.AddItem temp
            End If
        End If
    Next v
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'save searches back to registry
    Dim v As Integer
    For v = 0 To 9
        If Len(Trim(cboFind.List(v))) <> 0 Then
            SaveSetting "PSST SOFTWARE\" + App.Title, "FindText", "Text" + Str(v), cboFind.List(v)
        End If
        If Len(Trim(cboReplace.List(v))) <> 0 Then
            SaveSetting "PSST SOFTWARE\" + App.Title, "ReplaceText", "Text" + Str(v), cboReplace.List(v)
        End If
    Next v
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.EditEnable
End Sub
