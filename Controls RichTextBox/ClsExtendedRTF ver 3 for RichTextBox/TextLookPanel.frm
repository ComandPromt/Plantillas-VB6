VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form TextLookPanel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text Appearance Panel"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Format"
      Height          =   375
      Left            =   3960
      TabIndex        =   32
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox chkAccumulatestyles 
      Caption         =   "Accumulate Styles"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame FrmStandard 
      Caption         =   "Standard Properies"
      Height          =   3735
      Left            =   6240
      TabIndex        =   29
      Top             =   120
      Width           =   2655
      Begin VB.Label Label4 
         Caption         =   $"TextLookPanel.frx":0000
         Height          =   3255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame FrmUnderlines 
      Caption         =   "Underlines"
      Height          =   3735
      Left            =   6240
      TabIndex        =   27
      Top             =   120
      Width           =   2655
      Begin VB.Label Label3 
         Caption         =   $"TextLookPanel.frx":008C
         Height          =   3255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "show TextRTF"
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ripple Values"
      Height          =   3735
      Left            =   6240
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   1440
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
         Begin VB.CheckBox Check1 
            Caption         =   "Down 1st"
            Height          =   375
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Wave Len"
         Height          =   735
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   1095
         Begin MSComCtl2.UpDown UpDown3 
            Height          =   405
            Left            =   720
            TabIndex        =   21
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   714
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text4"
            BuddyDispid     =   196619
            OrigLeft        =   720
            OrigTop         =   240
            OrigRight       =   975
            OrigBottom      =   495
            Max             =   20
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Text            =   "1"
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Amplitude"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Text            =   "5"
            Top             =   240
            Width           =   585
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   720
            TabIndex        =   17
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "Text1"
            BuddyDispid     =   196621
            OrigLeft        =   7200
            OrigTop         =   1440
            OrigRight       =   7455
            OrigBottom      =   1815
            Max             =   40
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Start Value"
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Text            =   "0"
            Top             =   240
            Width           =   585
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   375
            Left            =   720
            TabIndex        =   14
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            BuddyControl    =   "Text2"
            BuddyDispid     =   196623
            OrigLeft        =   960
            OrigTop         =   360
            OrigRight       =   1215
            OrigBottom      =   735
            Max             =   40
            Min             =   -40
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
      End
      Begin VB.Label Label2 
         Caption         =   $"TextLookPanel.frx":0196
         Height          =   1335
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ransom Colour"
      Height          =   3735
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   1575
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton Option1 
            Caption         =   "Rnd Fore & Back"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   11
            Top             =   765
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rnd Back"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   10
            Top             =   510
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rnd Text"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   9
            Top             =   255
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "No Colour"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Remember that Ransom colours and text are random. The sample is not exactly what you will get."
            Height          =   1215
            Left            =   0
            TabIndex        =   22
            Top             =   1080
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample"
      Height          =   3735
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      Begin RichTextLib.RichTextBox DemoRTB 
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5953
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"TextLookPanel.frx":0252
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text3 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2280
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do it"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "TextLookPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2002 Roger Gilchrist
'rojgilkrist@hotmail.com
'very new; not much comment
'you'll have to work it out

Option Explicit
Private Demo As New ClsRTFFontPainter
Private Reset As Boolean
Private Accumulator As String

Private Sub ActivateTools(Clr As Boolean, Rpl As Boolean, Wav As Boolean, Undl As Boolean, Stnd As Boolean)

    FrmUnderlines.Visible = Undl
    FrmStandard.Visible = Stnd
    Frame4.Visible = Clr
    Frame5.Visible = Rpl
    Frame6.Visible = Wav
    Frame3.Visible = Wav
    Check1.Visible = Wav
    Select Case LCase$(List1.List(List1.ListIndex))
      Case "up", "down", "subscript", "superscript"
        DemoRTB.Find "ff", 1
      Case "visible"
        DemoRTB.Find "show ", 1
      Case Else
        DemoRTB.Find "This line will show the effect.", 1
    End Select

End Sub

Private Sub Check1_Click()

    If Reset = False Then
        List1_Click
    End If

End Sub

Private Sub Check2_Click()

    If Check2.Value Then
        Text3.Visible = True
        DemoRTB.Height = 1935
      Else 'CHECK2.VALUE = FALSE
        Text3.Visible = False
        DemoRTB.Height = 3375
    End If

End Sub

Private Sub Command1_Click(Index As Integer)

    If Index = 0 Then
        TakeAction False
    End If
    Unload TextLookPanel

End Sub

Private Sub Command2_Click()

    Demo.NoFormatting
    Accumulator = ""

End Sub

Private Sub DemoRTB_Change()

    Text3.Text = DemoRTB.TextRTF

End Sub

Private Sub Form_Load()

    Me.Height = 4875
    Demo.AssignControls DemoRTB, ExtendedRTFDemo.CommonDialog1

    'The name of this   V_________V needs to match that being used by you RichTextBox
    Command1(0).Enabled = RTBLooks.IsSelection
    Command1(0).Caption = IIf(Command1(0).Enabled, "Do It", "No Selection")
    DemoRTB.SelFontSize = 14
    DemoRTB.Text = "This line will not be touched." & vbNewLine & "This line will show the effect." & vbNewLine & "This line will not be touched."
    'List1 is Sorted=True as the rest of the form actually reads the lcase string value of this list
    'You can add either Lcase,Ucase or ProperCase strings here but make sure you use lcase everywhere else
    With List1
        .AddItem "Ransom"
        .AddItem "Ripple Baseline"
        .AddItem "Ripple Baseline1"
        .AddItem "Ripple Height"
        .AddItem "Up"
        .AddItem "Down"
        .AddItem "SubScript"
        .AddItem "SuperScript"
        .AddItem "ALLCAPS"
        .AddItem "Visible"
        .AddItem "Bold"
        .AddItem "Italic"
        .AddItem "Strikethru"
        .AddItem "Underline"
        .AddItem "UnderlineDot"
        .AddItem "UnderlineDash"
        .AddItem "UnderlineDashDot"
        .AddItem "UnderlineDashDotDot"
        .AddItem "UnderlineWave"
        .AddItem "UnderlineWord"
        .AddItem "UnderlineDouble"
        .AddItem "UnderlineHair"
        .AddItem "UnderlineThick"
    End With 'LIST1

End Sub

Private Sub Frame7_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub List1_Click()

  Static prevIndex As Integer

    If List1.ListIndex <> prevIndex Then
        Reset = True
        Text4.Text = 1
        text1.Text = 5
        Text2.Text = 0
        Check1.Value = vbUnchecked
        Reset = False
        prevIndex = List1.ListIndex
    End If
    If chkAccumulatestyles.Value = vbUnchecked Then
        Demo.NoFormatting
        Accumulator = LCase$(List1.List(List1.ListIndex))
      Else 'NOT CHKACCUMULATESTYLES.VALUE...
        Accumulator = Accumulator & "|" & LCase$(List1.List(List1.ListIndex))
    End If
    Select Case LCase$(List1.List(List1.ListIndex))
      Case "allcaps"
        ActivateTools False, False, False, False, False
      Case "up", "down"
        ActivateTools False, True, False, False, False
      Case "ransom"
        ActivateTools True, False, False, False, False
      Case "ripple baseline", "ripple baseline1"
        ActivateTools False, True, True, False, False
      Case "ripple height"
        ActivateTools False, True, True, False, False
      Case "visible"
        ActivateTools False, False, False, False, False
      Case "superscript", "subscript", "bold", "italic", "strikethru"
        ActivateTools False, False, False, False, True
      Case "underline", "underline", "underlinedot", "underlinedash", "underlinedashdot", "underlinedashdotdot", "underlinewave", "underlineword", "underlinedouble", "underlinehair", "underlinethick"
        ActivateTools False, False, False, True, False
    End Select
    TakeAction True

End Sub

Private Sub Option1_Click(Index As Integer)

    If Reset = False Then
        List1_Click
    End If

End Sub

Private Function Option_True(C As Variant) As Integer

  'returns index of option array member which is true
  'call SomeVariable=Option_True(SomeOptionArray())

  Dim ctl As Control

    For Each ctl In C
        If ctl.Value Then
            Option_True = ctl.Index
            Exit Function '>---> Bottom
        End If
    Next ctl

End Function

Private Sub TakeAction(DemoDoc As Boolean)

  Dim Range As Integer
  Dim Colourise As Integer
  Dim InitialDirection As Boolean
  Dim InitialStartValue As Integer
  Dim WaveVal As Integer
  Dim Target As Variant
  Dim PrevAccum As Variant
  Dim AccumArray As Variant, AccumPart As Variant

    Range = Val(text1.Text)
    InitialDirection = (Check1.Value = vbChecked)
    Colourise = Option_True(Option1)
    InitialStartValue = Val(Text2.Text)
    WaveVal = Val(Text4.Text)

    If DemoDoc Then
        Set Target = Demo
        AccumArray = Split(LCase$(List1.List(List1.ListIndex)), "|") 'only do selection if demo
      Else 'DEMODOC = FALSE
        Set Target = RTBLooks
        AccumArray = Demo.FontLookArray(Accumulator)
        Beep

    End If

    For Each AccumPart In AccumArray
        If PrevAccum <> AccumPart Then
            PrevAccum = AccumPart
            Select Case LCase$(AccumPart)

              Case "allcaps"
                Target.SelCaps = Not Target.SelCaps
              Case "up"
                Target.SelUp = Range
              Case "down"
                Target.SelDown = Range
              Case "superscript"
                Target.SelSuper = Not Target.SelSuper
              Case "subscript"
                Target.SelSub = Not Target.SelSub
              Case "ransom"
                Target.Ransom Colourise
              Case "ripple baseline"
                Target.RippleEngine BaseLine, Range, InitialDirection, InitialStartValue, WaveVal
              Case "ripple baseline1"
                Target.RippleEngine BaseLine1, Range, InitialDirection, InitialStartValue, WaveVal
              Case "ripple height"
                Target.RippleEngine THeight, Range, InitialDirection, InitialStartValue, WaveVal
              Case "visible"
                Target.SelVisible = Not Target.SelVisible
              Case "underline"
                Target.SelUnderline = Not Target.SelUnderline
              Case "bold"
                Target.SelBold = Not Target.SelBold
              Case "italic"
                Target.SelItalic = Not Target.SelItalic
              Case "strikethru"
                Target.SelStrikeThru = Not Target.SelStrikeThru
              Case "underlinedot"
                Target.SelDot = Not Target.SelDot
              Case "underlinedash"
                Target.SelDash = Not Target.SelDash
              Case "underlinedashdot"
                Target.SelDashd = Not Target.SelDashd
              Case "underlinedashdotdot"
                Target.SelDashdd = Not Target.SelDashdd
              Case "underlinewave"
                Target.SelWave = Not Target.SelWave
              Case "underlineword"
                Target.SelUlWord = Not Target.SelUlWord
              Case "underlinedouble"
                Target.SelUlDouble = Not Target.SelUlDouble
              Case "underlinehair"
                Target.SelHair = Not Target.SelHair
              Case "underlinethick"
                Target.SelThick = Not Target.SelThick
            End Select
        End If
    Next AccumPart

End Sub

Private Sub Text1_Change()

    If Reset = False Then
        List1_Click
    End If

End Sub

Private Sub Text2_Change()

    If Reset = False Then
        List1_Click
    End If

End Sub

Private Sub Text4_Change()

    If Reset = False Then
        List1_Click
    End If

End Sub

':) Ulli's VB Code Formatter V2.13.6 (28/08/2002 2:38:47 PM) 8 + 281 = 289 Lines
