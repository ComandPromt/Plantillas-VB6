VERSION 5.00
Begin VB.Form FTestSplitter 
   Caption         =   "Test Splitters"
   ClientHeight    =   5610
   ClientLeft      =   2655
   ClientTop       =   2955
   ClientWidth     =   7530
   ClipControls    =   0   'False
   DrawStyle       =   6  'Inside Solid
   LinkTopic       =   "Form1"
   MouseIcon       =   "Splitter.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5610
   ScaleWidth      =   7530
   Begin VB.TextBox txtB 
      Height          =   1260
      Left            =   3990
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Splitter.frx":0212
      Top             =   1860
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.TextBox txtA 
      Height          =   1260
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Splitter.frx":022E
      Top             =   1995
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.PictureBox pbBack 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   3660
      Left            =   900
      Picture         =   "Splitter.frx":025C
      ScaleHeight     =   3600
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.Image imgHorizontal 
      Height          =   480
      Left            =   6240
      MouseIcon       =   "Splitter.frx":98DE
      Picture         =   "Splitter.frx":9A30
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgVertical 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "Splitter.frx":9B82
      Picture         =   "Splitter.frx":9CD4
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FTestSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vsplit As New CVSplitter
Private hsplit As New CHSplitter
Private ctlA  As Control, ctlB As Control
Private ordControl As Integer

Private Sub cboSplit_Click()
    NewSplit
End Sub

Private Sub Form_Load()
    Show
    NewSplit
End Sub

Sub NewSplit()
    Set vsplit = Nothing
    Set hsplit = Nothing

    If Not ctlA Is Nothing Then ctlA.Visible = False
    If Not ctlB Is Nothing Then ctlB.Visible = False
        Set ctlA = txtA
        Set ctlB = txtB
        pbBack.Visible = False
        Set ctlA.Container = Me
        Set ctlB.Container = Me
        ctlA.Left = Width * 0.05
        ctlA.Top = Height * 0.05
        ctlA.Width = Width * 0.6
        ctlA.Height = Height * 0.6
        ctlB.Left = ctlA.Width + 1
        ctlB.Top = ctlA.Height + 1
        ctlA.Visible = True
        ctlB.Visible = True
        ctlA.Container.MouseIcon = imgHorizontal.Picture
 
            If Not hsplit.Create(vTopControl:=ctlA, _
                                vBottomControl:=ctlB, _
                                vBorderPixels:=4, _
                                vAutoBorder:=False, _
                                vResizeable:=True) Then
                MsgBox "Can't create splitter"
                End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsplit.VSplitter_MouseDown Button, Shift, X, Y
        hsplit.HSplitter_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsplit.VSplitter_MouseMove Button, Shift, X, Y
        hsplit.HSplitter_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsplit.VSplitter_MouseUp Button, Shift, X, Y
        hsplit.HSplitter_MouseUp Button, Shift, X, Y
End Sub
Private Sub Form_Resize()
        vsplit.VSplitter_Resize
        hsplit.HSplitter_Resize
End Sub

Private Sub txtA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsplit.VSplitter_MouseOff
        hsplit.HSplitter_MouseOff
End Sub

Private Sub txtB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsplit.VSplitter_MouseOff
        hsplit.HSplitter_MouseOff
End Sub



