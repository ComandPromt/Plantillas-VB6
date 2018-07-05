VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Slider Bar Example"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctHorizontalSlider 
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   4410
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   1575
      Width           =   4755
   End
   Begin VB.PictureBox pctVerticalSlider 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   4095
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4575
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   105
      Width           =   135
   End
   Begin VB.PictureBox pctHorizontalShadow 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   4410
      ScaleHeight     =   60
      ScaleWidth      =   4650
      TabIndex        =   5
      Top             =   1470
      Visible         =   0   'False
      Width           =   4650
   End
   Begin VB.PictureBox pctVerticalShadow 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   4200
      ScaleHeight     =   4575
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox fraRight 
      BackColor       =   &H80000005&
      Height          =   1170
      Left            =   4410
      ScaleHeight     =   1110
      ScaleWidth      =   4680
      TabIndex        =   3
      Top             =   105
      Width           =   4740
      Begin VB.Label lblFrameTop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Top Frame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3270
      End
   End
   Begin VB.PictureBox fraLeft 
      BackColor       =   &H80000005&
      Height          =   4635
      Left            =   105
      ScaleHeight     =   4575
      ScaleWidth      =   3840
      TabIndex        =   2
      Top             =   105
      Width           =   3900
      Begin VB.Label lblFrameLeft 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Left (""menu"") Frame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   0
         TabIndex        =   9
         Top             =   2205
         Width           =   3270
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sizable Frame Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1800
         Left            =   210
         TabIndex        =   7
         Top             =   210
         Width           =   3270
      End
   End
   Begin VB.PictureBox fraBottom 
      BackColor       =   &H80000005&
      Height          =   2535
      Left            =   4410
      ScaleHeight     =   2475
      ScaleWidth      =   4680
      TabIndex        =   4
      Top             =   1785
      Width           =   4740
      Begin VB.Label lblFrameBottom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom Frame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Variables to toggle frame resize modes

'   I made this form based on an example I found on www.freevbcode.com.
'   The code is completely different, but the concept is the same. I have
'   looked for a good example of how to make an "Outlook" style split form
'   in VB without doing it with an .ocx.

'   Most of the examples I found for making splitter bars were
'   very limited in that they did not offer good resizing capablilty. This
'   form uses picture boxes as "frames" since they are container objects and
'   fire a resize event when they are resized. Putting resizing code for each
'   frame in the Resize event for that frame makes it much easier to maintain
'   and code.

'   The down side? Picture boxes have a good bit of overhead. Not tragically so,
'   but they are much heavier than a text box. It is worth it to me though because
'   I can build each "frame" independent of each other.

'   I have included some very basic examples of using the resize event for each
'   frame, but in the interest of keeping it simple, didn't put much. You can add
'   just about any control to a frame that you can add to a form, so you should
'   be able to really create some nice UI's with this template.




Dim blnDragging As Boolean

Private Sub Form_Load()

'   Set drag bar and shadow widths
    pctVerticalShadow.Width = 100
    pctHorizontalSlider.Height = 50
    pctVerticalSlider.Width = 75

End Sub

Private Sub Form_Resize()

'   Make sure window is not minimized and form is not too small
If Not WindowState = vbMinimized Then
    If Height < 1200 Or Width < 1200 Then
        Exit Sub
    Else
        DrawFrames
        DrawDragBars
    End If
End If

End Sub


Private Sub fraBottom_Resize()
    
    lblFrameBottom.Width = fraBottom.Width
    lblFrameBottom.Top = (fraBottom.Height / 2) - (lblFrameBottom.Height / 2)

End Sub

Private Sub fraLeft_Resize()
    
    '   Put any control resize code for this frame here
    lblTitle.Width = fraLeft.Width * 0.8
    lblFrameLeft.Width = fraLeft.Width
    

End Sub

Private Sub fraRight_Resize()
    lblFrameTop.Width = fraRight.Width
    lblFrameTop.Top = (fraRight.Height / 2) - (lblFrameTop.Height / 2)
End Sub

Private Sub pctHorizontalSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '   Turn on drag mode
    blnDragging = True
    
End Sub

Private Sub pctHorizontalSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '   If Dragging, show shadow at mouse location
    If blnDragging Then
    
        If Not pctHorizontalShadow.Visible Then pctHorizontalShadow.Visible = True
        pctHorizontalShadow.Top = pctHorizontalSlider.Top + Y
    
    End If

End Sub

Private Sub pctHorizontalSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '   User let go of mouse button...redraw using current location
    blnDragging = False
    pctHorizontalShadow.Visible = False
    
    DrawFrames

End Sub

Private Sub pctVerticalSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dragging has started
blnDragging = True

End Sub


Private Sub pctVerticalSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If blnDragging Then

    pctVerticalShadow.Visible = True
    pctVerticalShadow.Left = pctVerticalSlider.Left + X

End If

End Sub


Private Sub pctVerticalSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '   Stop dragging, draw frames based on current location
    blnDragging = False
    pctVerticalShadow.Visible = False
    DrawFrames

End Sub



Private Sub DrawFrames()

'       Draw frames based on the current drag position

    If pctHorizontalShadow.Top + 300 > ScaleHeight Then
        pctHorizontalShadow.Top = ScaleHeight - 300
    End If
    
    If pctHorizontalShadow.Top - 300 < 0 Then
        pctHorizontalShadow.Top = 300
    End If
    
    If pctVerticalShadow.Left + 300 > ScaleWidth Then pctVerticalShadow.Left = ScaleWidth - 300
    If pctVerticalShadow.Left - 300 < 0 Then pctVerticalShadow.Left = 300
    
    pctVerticalSlider.Left = pctVerticalShadow.Left
    
    pctHorizontalSlider.Top = pctHorizontalShadow.Top
    
    fraLeft.Height = Height - fraLeft.Top - 400
    fraLeft.Width = pctVerticalSlider.Left - fraLeft.Left
    
    With fraRight
        .Left = pctVerticalSlider.Left + pctVerticalSlider.Width
        .Width = (Width - pctVerticalSlider.Left)
        .Height = pctHorizontalSlider.Top - fraRight.Top
    End With
    
    fraBottom.Left = fraRight.Left
    fraBottom.Width = fraRight.Width
    
    fraBottom.Top = pctHorizontalSlider.Top + pctHorizontalSlider.Height
    fraBottom.Height = (fraLeft.Height - pctHorizontalSlider.Top)
    
    DrawDragBars

End Sub

Private Sub DrawDragBars()
        
        '   Resize drag bars to correct deminsions since frames were resized

        pctVerticalSlider.Height = fraLeft.Height
        pctVerticalShadow.Height = pctVerticalSlider.Height
        
        pctHorizontalShadow.Left = fraBottom.Left
        pctHorizontalSlider.Left = fraBottom.Left
        
        pctHorizontalShadow.Width = fraBottom.Width
        pctHorizontalSlider.Width = fraBottom.Width

End Sub
