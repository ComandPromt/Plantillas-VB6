VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[Esc | Space | Enter|  mouse touch to continue"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Silly As New ClsRTFFontPainter

'Form has been locked because it is so sensitive to settings
Private Sub Command1_Click()

  'this button is off screen but set to Cancel=True and Default=True
  'and top of TabIndex so gets focus when form appears
  ' so any Esc, SpaceBar and Enter keys trigger it

    Unload Form1

End Sub

Private Sub Form_Load()

  'RichTextBox1 TabStop=False | Locked=True | Enabled=False | HideSelection=True
  'reason:      Loses focus   | Can't Edit  | Can't Touch   | Selection reverse colouring disappears when loses focus
  'The spaces in the next string are for layout purposes

    RichTextBox1.Text = "Welcome! This demo program is  based on manipulating RTF code!"
    'select whole Text
    RichTextBox1.SelLength = Len(RichTextBox1.Text)
    'set control to class            'just because class needs it
    Silly.AssignControls RichTextBox1, ExtendedRTFDemo.CommonDialog1
    Timer1.Enabled = True 'start timer

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Command1_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Timer1.Enabled = False 'turn off timer
    Unload Form1

End Sub

Private Sub Timer1_Timer()

  Static Flicker As Boolean ' trigger for animation effect

    Flicker = Not Flicker
    Silly.RippleEngine BaseLine, 8, Flicker, 3, 3
    Silly.SpectrumSector s3GreenCyan, Flicker, Flicker, False
    'due to time considerations of the code manipulation this text length and
    'Ripple settings are near the maximum you can use for this sort of animation

End Sub

':) Ulli's VB Code Formatter V2.13.6 (26/08/2002 4:37:37 PM) 2 + 53 = 55 Lines
