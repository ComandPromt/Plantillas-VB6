VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   3150
   ClientTop       =   2655
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5610
   ScaleWidth      =   6015
   Begin VB.HScrollBar HBar 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2850
      Width           =   3150
   End
   Begin VB.VScrollBar VBar 
      Height          =   2535
      Left            =   2850
      TabIndex        =   2
      Top             =   45
      Width           =   255
   End
   Begin VB.PictureBox OuterPict 
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   0
      Width           =   2835
      Begin VB.PictureBox InnerPict 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7200
         Left            =   0
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   7200
         ScaleWidth      =   9600
         TabIndex        =   1
         Top             =   0
         Width           =   9600
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   525
            Left            =   1350
            TabIndex        =   4
            Top             =   2760
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetScrollBars()
    ' Set scroll bar properties.
    VBar.Min = 0
    VBar.Max = OuterPict.ScaleHeight - InnerPict.Height
    VBar.LargeChange = OuterPict.ScaleHeight
    VBar.SmallChange = OuterPict.ScaleHeight / 5
    
    HBar.Min = 0
    HBar.Max = OuterPict.ScaleWidth - InnerPict.Width
    HBar.LargeChange = OuterPict.ScaleWidth
    HBar.SmallChange = OuterPict.ScaleWidth / 5
End Sub
Private Sub Form_Resize()
Dim got_wid As Single
Dim got_hgt As Single
Dim need_wid As Single
Dim need_hgt As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    If WindowState = vbMinimized Then Exit Sub

    need_wid = InnerPict.Width + (OuterPict.Width - OuterPict.ScaleWidth)
    need_hgt = InnerPict.Height + (OuterPict.Height - OuterPict.ScaleHeight)
    got_wid = ScaleWidth
    got_hgt = ScaleHeight

    ' See which scroll bars we need.
    need_hbar = (need_wid > got_wid)
    If need_hbar Then got_hgt = got_hgt - HBar.Height

    need_vbar = (need_hgt > got_hgt)
    If need_vbar Then
        got_wid = got_wid - VBar.Width
        If Not need_hbar Then
            need_hbar = (need_wid > got_wid)
            If need_hbar Then got_hgt = got_hgt - HBar.Height
        End If
    End If

    OuterPict.Move 0, 0, got_wid, got_hgt

    If need_hbar Then
        HBar.Move 0, got_hgt, got_wid
        HBar.Visible = True
    Else
        HBar.Visible = False
    End If

    If need_vbar Then
        VBar.Move got_wid, 0, VBar.Width, got_hgt
        VBar.Visible = True
    Else
        VBar.Visible = False
    End If
    
    SetScrollBars
End Sub

Private Sub HBar_Change()
    InnerPict.Left = HBar.Value
End Sub


Private Sub HBar_Scroll()
    InnerPict.Left = HBar.Value
End Sub


Private Sub VBar_Change()
    InnerPict.Top = VBar.Value
End Sub


Private Sub VBar_Scroll()
    InnerPict.Top = VBar.Value
End Sub


