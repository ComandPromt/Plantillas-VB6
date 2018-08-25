VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAnimator 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprite Animator"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRight 
      Height          =   435
      Left            =   4020
      Picture         =   "frmAnimator.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4500
      Width           =   495
   End
   Begin VB.CommandButton cmdLeft 
      Height          =   435
      Left            =   3360
      Picture         =   "frmAnimator.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4500
      Width           =   495
   End
   Begin VB.HScrollBar hscTimer 
      Height          =   195
      LargeChange     =   100
      Left            =   6660
      Max             =   1000
      Min             =   10
      SmallChange     =   10
      TabIndex        =   16
      Top             =   4080
      Value           =   50
      Width           =   1035
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "&Animate"
      Height          =   495
      Left            =   6060
      TabIndex        =   15
      Top             =   4440
      Width           =   1695
   End
   Begin VB.HScrollBar hscRows 
      Height          =   195
      LargeChange     =   10
      Left            =   6660
      Max             =   32
      Min             =   1
      TabIndex        =   14
      Top             =   3780
      Value           =   4
      Width           =   1035
   End
   Begin VB.HScrollBar hscColumns 
      Height          =   195
      LargeChange     =   10
      Left            =   6660
      Max             =   32
      Min             =   1
      TabIndex        =   9
      Top             =   3480
      Value           =   8
      Width           =   1035
   End
   Begin VB.HScrollBar hscHeight 
      Height          =   195
      LargeChange     =   8
      Left            =   6660
      Max             =   128
      Min             =   8
      TabIndex        =   8
      Top             =   3180
      Value           =   63
      Width           =   1035
   End
   Begin VB.HScrollBar hscWidth 
      Height          =   195
      LargeChange     =   8
      Left            =   6660
      Max             =   128
      Min             =   8
      TabIndex        =   7
      Top             =   2880
      Value           =   63
      Width           =   1035
   End
   Begin VB.PictureBox picSprite 
      BackColor       =   &H00C0C0C0&
      Height          =   2250
      Left            =   5160
      ScaleHeight     =   146
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   166
      TabIndex        =   2
      Top             =   60
      Width           =   2550
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   60
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load Image"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1260
      Left            =   180
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Frame:"
      Height          =   195
      Index           =   5
      Left            =   5100
      TabIndex        =   22
      Top             =   2580
      Width           =   480
   End
   Begin VB.Label lblFrame 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   5880
      TabIndex        =   21
      Top             =   2580
      Width           =   90
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Caption         =   "200"
      Height          =   195
      Left            =   5880
      TabIndex        =   18
      Top             =   4080
      Width           =   270
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Delay:"
      Height          =   195
      Index           =   4
      Left            =   5100
      TabIndex        =   17
      Top             =   4080
      Width           =   450
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Rows:"
      Height          =   195
      Index           =   3
      Left            =   5100
      TabIndex        =   13
      Top             =   3780
      Width           =   450
   End
   Begin VB.Label lblRows 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   5880
      TabIndex        =   12
      Top             =   3780
      Width           =   90
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Columns:"
      Height          =   195
      Index           =   2
      Left            =   5100
      TabIndex        =   11
      Top             =   3480
      Width           =   645
   End
   Begin VB.Label lblColumns 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   3480
      Width           =   90
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      Caption         =   "64"
      Height          =   195
      Left            =   5880
      TabIndex        =   6
      Top             =   3180
      Width           =   180
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "64"
      Height          =   195
      Left            =   5880
      TabIndex        =   5
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   195
      Index           =   1
      Left            =   5100
      TabIndex        =   4
      Top             =   3180
      Width           =   510
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   195
      Index           =   0
      Left            =   5100
      TabIndex        =   3
      Top             =   2880
      Width           =   465
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3900
      Left            =   60
      Stretch         =   -1  'True
      Top             =   60
      Width           =   4800
   End
End
Attribute VB_Name = "frmAnimator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ret&, n&, curX&, curY&, curFrameX&, curFrameY&, curFrame&
Dim sWidth&, sHeight&, sTimer&, sRows&, sCols&
Dim sAnimate As Boolean

Private Sub cmdAnimate_Click()
    sAnimate = Not sAnimate
End Sub

Private Sub cmdLoad_Click()
    CD1.Filter = "Image Files|*.bmp;*.gif;*.jpg"
    CD1.DialogTitle = "Load Sprite Image File"
    CD1.InitDir = App.Path
    CD1.ShowOpen
    If Len(CD1.FileName) > 0 Then
        Load_File CD1.FileName
    End If
End Sub

Public Sub Load_File(ByVal fn$)
    Image1.Picture = LoadPicture(fn$)
    Picture1.Picture = Image1.Picture
    Update_Sprite
End Sub

Private Sub Form_Load()
    curFrame = 0
    curFrameX = 0
    curFrameY = 0
    sWidth = hscWidth.value
    lblWidth.Caption = sWidth
    sHeight = hscHeight.value
    lblHeight.Caption = sHeight
    sCols = hscColumns.value
    lblColumns.Caption = sCols
    sRows = hscRows.value
    lblRows.Caption = sRows
    sTimer = hscTimer.value
    lblTimer.Caption = hscTimer.value
    sAnimate = False
    Me.Show
    DoEvents
    Animation_Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub hscColumns_Change()
    sCols = hscColumns.value
    lblColumns.Caption = sCols
End Sub

Private Sub hscHeight_Change()
    sHeight = hscHeight.value
    lblHeight.Caption = sHeight
    Update_Sprite
End Sub

Private Sub hscRows_Change()
    sRows = hscRows.value
    lblRows.Caption = sRows
End Sub

Private Sub hscTimer_Change()
    sTimer = hscTimer.value
    lblTimer.Caption = sTimer
End Sub

Private Sub hscWidth_Change()
    sWidth = hscWidth.value
    lblWidth.Caption = sWidth
    Update_Sprite
End Sub

Public Sub Update_Sprite()
    curX = curFrameX * sWidth + 1
    curY = curFrameY * sHeight + 1
    ret = StretchBlt(picSprite.hdc, 0, 0, sWidth, sHeight, Picture1.hdc, 1 + curX, 1 + curY, sWidth, sHeight, SRCCOPY)
End Sub

Private Sub delay(ByVal ms&)
    Dim start&
    start = GetTickCount
    Do Until GetTickCount - start > ms
    Loop
End Sub

Private Sub Animation_Loop()
    Do Until False
        If sAnimate Then
            Move_Frame 1
            Update_Sprite
            delay sTimer
        End If
        DoEvents
    Loop
End Sub

Private Sub cmdRight_Click()
    Move_Frame 1
    Update_Sprite
End Sub

Private Sub cmdLeft_Click()
    Move_Frame -1
    Update_Sprite
End Sub

Public Sub Move_Frame(ByVal dir&)
    curFrame = curFrame + dir
    If curFrame > (sRows * sCols) - 1 Then curFrame = 0
    If curFrame < 0 Then curFrame = (sRows * sCols) - 1
    curFrameX = curFrame Mod sCols
    curFrameY = curFrame \ sCols
    lblFrame.Caption = curFrame
End Sub
