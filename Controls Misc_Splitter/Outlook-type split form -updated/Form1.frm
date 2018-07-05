VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MousePointer    =   9  'Size W E
   ScaleHeight     =   7170
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRightPane 
      Height          =   6975
      Left            =   2160
      MousePointer    =   1  'Arrow
      ScaleHeight     =   6915
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.PictureBox picLeftPane 
      BackColor       =   &H8000000C&
      CausesValidation=   0   'False
      Height          =   6975
      Left            =   120
      MousePointer    =   1  'Arrow
      ScaleHeight     =   6915
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.PictureBox picInnerFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   4215
         ScaleWidth      =   1455
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
         Begin VB.CommandButton cmdScrollDown 
            Height          =   255
            Left            =   1080
            Picture         =   "Form1.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2880
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdScrollUp 
            Height          =   255
            Left            =   1080
            Picture         =   "Form1.frx":00A2
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   2
            Left            =   480
            Picture         =   "Form1.frx":0144
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   9
            Top             =   2760
            Width           =   480
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   1
            Left            =   480
            Picture         =   "Form1.frx":0586
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   8
            Top             =   1560
            Width           =   480
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   480
            Picture         =   "Form1.frx":09C8
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   7
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   12
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   11
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   10
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Command4"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Command3"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Command2"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Command1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SPLITTER_WIDTH = 50
Private Const BTN_HEIGHT = 315
Private Const SCROLL_UP = -BTN_HEIGHT
Private Const SCROLL_DOWN = BTN_HEIGHT
Private Const DRAW_HIDDEN = 0
Private Const DRAW_RAISED = 1
Private Const DRAW_INSET = 2
Private Const BOX_BORDER = 50
Private Const PIC_OFFSET = 300
Private Const PIC_SPACING = 700
Private Const LABEL_SPACING = 100

Private sglLeftPane As Single
Private bIsDragging As Boolean
Private strPics() As String
Private strLabels() As String
Private iCurButton As Integer
Private iFirstVis As Integer

Private Sub ArrangeControls()
Dim i As Integer
Dim sglRightPane As Single

    If WindowState = vbMinimized Then Exit Sub
    picLeftPane.Move 0, 0, sglLeftPane, ScaleHeight
    sglRightPane = (ScaleWidth - SPLITTER_WIDTH) - sglLeftPane
    If sglRightPane < 0 Then sglRightPane = 0
    picRightPane.Move (sglLeftPane + SPLITTER_WIDTH), 0, sglRightPane, ScaleHeight
    cmdButton(0).Move picLeftPane.ScaleLeft, picLeftPane.ScaleTop, picLeftPane.ScaleWidth, BTN_HEIGHT
    If cmdButton.Count > 1 Then
        For i = 1 To cmdButton.Count - 1
            If cmdButton(i).Tag = "TOP" Then
                cmdButton(i).Move picLeftPane.ScaleLeft, picLeftPane.ScaleTop + (BTN_HEIGHT * i), picLeftPane.ScaleWidth, BTN_HEIGHT
                iCurButton = i
            Else
                cmdButton(i).Move picLeftPane.ScaleLeft, picLeftPane.ScaleHeight - (BTN_HEIGHT * (cmdButton.Count - i)), picLeftPane.ScaleWidth, BTN_HEIGHT
            End If
        Next
    End If
    DrawInnerFrame
    DrawPics
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim i As Integer
Dim iStart As Integer
Dim iEnd As Integer
Dim iStep As Integer
Dim iDir As Integer

    If cmdButton(Index).Tag = "TOP" Then
        iStart = Index + 1
        iEnd = cmdButton.Count - 1
        iStep = 1
        iDir = SCROLL_DOWN
    Else
        iStart = Index
        iEnd = 1
        iStep = -1
        iDir = SCROLL_UP
    End If
    For i = iStart To iEnd Step iStep
        ScrollBtn i, iDir
    Next
    If iCurButton <> Index Then
        iCurButton = Index
        iFirstVis = 0
        cmdScrollUp.Visible = False
        DrawInnerFrame
        DrawPics
    End If
    picLeftPane.SetFocus
End Sub

Private Sub cmdScrollDown_Click()
    iFirstVis = iFirstVis + 1
    If iFirstVis >= (Picture1.Count - 1) Then iFirstVis = (Picture1.Count - 1)
    If Not Picture1(iFirstVis).Visible Then
        iFirstVis = iFirstVis - 1
    Else
        DrawPics
    End If
    If iFirstVis > 0 Then cmdScrollUp.Visible = True
End Sub

Private Sub cmdScrollUp_Click()
    iFirstVis = iFirstVis - 1
    If iFirstVis <= 0 Then
        iFirstVis = 0
        cmdScrollUp.Visible = False
    End If
    DrawPics
End Sub

Private Sub Form_Load()
Dim i As Integer

    cmdButton(0).Tag = "TOP"
    For i = 1 To cmdButton.Count - 1
        cmdButton(i).Tag = "BOTTOM"
    Next
    sglLeftPane = 2000
    
    ReDim strPics(cmdButton.Count, Picture1.Count)
    strPics(0, 0) = "CDROM"
    strPics(0, 1) = "ENVELOPE"
    strPics(0, 2) = "NOTEPAD"
    strPics(1, 0) = "KEY"
    strPics(1, 1) = "POSTIT"
    strPics(1, 2) = ""
    strPics(2, 0) = "NOTEPAD"
    strPics(2, 1) = ""
    strPics(2, 2) = ""
    strPics(3, 0) = "NOTEPAD"
    strPics(3, 1) = "POSTIT"
    strPics(3, 2) = "KEY"
    
    ReDim strLabels(cmdButton.Count, Label1.Count)
    strLabels(0, 0) = "CDROM"
    strLabels(0, 1) = "ENVELOPE"
    strLabels(0, 2) = "NOTEPAD"
    strLabels(1, 0) = "KEY"
    strLabels(1, 1) = "POSTIT"
    strLabels(1, 2) = ""
    strLabels(2, 0) = "NOTEPAD"
    strLabels(2, 1) = ""
    strLabels(2, 2) = ""
    strLabels(3, 0) = "NOTEPAD"
    strLabels(3, 1) = "POSTIT"
    strLabels(3, 2) = "KEY"
    
    iCurButton = 0
    iFirstVis = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bIsDragging = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bIsDragging Then
        sglLeftPane = X
        If sglLeftPane < 0 Then sglLeftPane = 0
        If sglLeftPane > ScaleWidth Then sglLeftPane = ScaleWidth - SPLITTER_WIDTH
        ArrangeControls
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bIsDragging = False
End Sub

Private Sub Form_Resize()
    ArrangeControls
End Sub

Private Sub picInnerFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

    If Picture1.Count > 0 Then
        For i = 0 To Picture1.Count - 1
            Make3D Picture1(i), DRAW_HIDDEN
        Next
    End If
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Make3D Picture1(Index), DRAW_INSET
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Make3D Picture1(Index), DRAW_RAISED
End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Make3D Picture1(Index), DRAW_RAISED
End Sub

Private Sub ScrollBtn(iBtnIndex As Integer, iDir As Integer)
Dim i As Integer
Dim iBtnStep As Integer
Dim iEndPos As Integer

    If iDir = SCROLL_UP Then
        iEndPos = picLeftPane.ScaleTop + (BTN_HEIGHT * iBtnIndex)
        cmdButton(iBtnIndex).Tag = "TOP"
    Else
        iEndPos = picLeftPane.ScaleHeight - (BTN_HEIGHT * (cmdButton.Count - iBtnIndex))
        cmdButton(iBtnIndex).Tag = "BOTTOM"
    End If
    For i = cmdButton(iBtnIndex).Top To iEndPos Step iDir
        cmdButton(iBtnIndex).Move picLeftPane.ScaleLeft, i, picLeftPane.ScaleWidth, BTN_HEIGHT
    Next
    If i <> iEndPos Then cmdButton(iBtnIndex).Move picLeftPane.ScaleLeft, iEndPos, picLeftPane.ScaleWidth, BTN_HEIGHT
End Sub

Private Sub Make3D(ctl As Control, iMode As Integer)
Dim lTopLeftCol As Long
Dim lBotRightCol As Long

    Select Case iMode
    Case DRAW_INSET
        lTopLeftCol = vbBlack
        lBotRightCol = vbWhite
    Case DRAW_RAISED
        lTopLeftCol = vbWhite
        lBotRightCol = vbBlack
    Case Else
        lTopLeftCol = picLeftPane.BackColor
        lBotRightCol = picLeftPane.BackColor
    End Select
    picInnerFrame.CurrentX = ctl.Left - BOX_BORDER
    picInnerFrame.CurrentY = ctl.Top - BOX_BORDER
    'left
    picInnerFrame.Line -(ctl.Left - BOX_BORDER, ctl.Top + ctl.Height + BOX_BORDER), lTopLeftCol
    'bottom
    picInnerFrame.Line -(ctl.Left + ctl.Width + BOX_BORDER, ctl.Top + ctl.Height + BOX_BORDER), lBotRightCol
    'right
    picInnerFrame.Line -(ctl.Left + ctl.Width + BOX_BORDER, ctl.Top - BOX_BORDER), lBotRightCol
    'top
    picInnerFrame.Line -(ctl.Left - BOX_BORDER, ctl.Top - BOX_BORDER), lTopLeftCol
End Sub

Private Sub DrawPics()
Dim i As Integer
Dim iLastVis As Integer

    If WindowState = vbMinimized Then Exit Sub
    If Picture1.Count > 0 Then
        For i = 0 To Picture1.Count - 1
            If Not bIsDragging Then
                If i >= iFirstVis And strPics(iCurButton, i) <> "" Then
                    Picture1(i).Picture = LoadResPicture(strPics(iCurButton, i), vbResIcon)
                    Picture1(i).Visible = True
                    Label1(i).Caption = strLabels(iCurButton, i)
                    Label1(i).Visible = True
                    iLastVis = i
                Else
                    Picture1(i).Visible = False
                    Label1(i).Visible = False
                End If
            End If
            If i = iFirstVis Then
                Picture1(i).Move (picInnerFrame.ScaleWidth - Picture1(i).Width) / 2, picInnerFrame.ScaleTop + PIC_OFFSET
            ElseIf i > iFirstVis Then
                Picture1(i).Move (picInnerFrame.ScaleWidth - Picture1(i).Width) / 2, Picture1(i - 1).Top + Picture1(i - 1).Height + PIC_SPACING
            End If
            Label1(i).Move (picInnerFrame.ScaleWidth - Label1(i).Width) / 2, Picture1(i).Top + Picture1(i).Height + LABEL_SPACING
        Next
        If iLastVis > iFirstVis And (Label1(iLastVis).Top + Label1(iLastVis).Height) > picInnerFrame.ScaleHeight Then
            cmdScrollDown.Visible = True
        Else
            cmdScrollDown.Visible = False
        End If
    End If
End Sub

Private Sub DrawInnerFrame()
    If iCurButton < (cmdButton.Count - 1) Then
        If (cmdButton(iCurButton + 1).Top - cmdButton(iCurButton).Top - cmdButton(iCurButton).Height) > 0 Then picInnerFrame.Move picLeftPane.ScaleLeft, cmdButton(iCurButton).Top + cmdButton(iCurButton).Height, picLeftPane.ScaleWidth, cmdButton(iCurButton + 1).Top - cmdButton(iCurButton).Top - cmdButton(iCurButton).Height
    Else
        If (picLeftPane.ScaleHeight - cmdButton(iCurButton).Top - cmdButton(iCurButton).Height) > 0 Then picInnerFrame.Move picLeftPane.ScaleLeft, cmdButton(iCurButton).Top + cmdButton(iCurButton).Height, picLeftPane.ScaleWidth, picLeftPane.ScaleHeight - cmdButton(iCurButton).Top - cmdButton(iCurButton).Height
    End If
    cmdScrollUp.Move picInnerFrame.ScaleWidth - cmdScrollUp.Width - 100, picInnerFrame.ScaleTop + 100
    cmdScrollDown.Move picInnerFrame.ScaleWidth - cmdScrollDown.Width - 100, picInnerFrame.ScaleHeight - cmdScrollDown.Height - 100
End Sub
