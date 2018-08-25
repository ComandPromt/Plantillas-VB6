VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extract Icons"
   ClientHeight    =   3735
   ClientLeft      =   5850
   ClientTop       =   2400
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3810
   Begin MSComctlLib.Toolbar tbLarge 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbSmall 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   14
      Top             =   630
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Icons"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.PictureBox picSmall 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1320
         ScaleHeight     =   240
         ScaleMode       =   0  'User
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox picLarge 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   360
         ScaleHeight     =   495
         ScaleMode       =   0  'User
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblIcon 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Current Icon Index:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblIcons 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of  Icons:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Small"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Large"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   3240
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   1440
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FilterIndex     =   1
   End
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   960
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   840
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.Label lblFile 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   3495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' This example demonstrates how to:
'   Extract both large and small icons from executables and dll's.
'   Draw them to a control with a device context handle (.hdc) such as a PictureBox.
'   Draw them to a control without an .hdc property such as an ImageList.
'   Dynamically populate both ImageList and ToolBar controls.
'
Dim glLargeIcons() As Long
Dim glSmallIcons() As Long
Dim lIndex         As Long
Dim lIcons         As Long
Dim sExeName       As String

Const LARGE_ICON As Integer = 32
Const SMALL_ICON As Integer = 16
Const DI_NORMAL = 3
Private Declare Function DrawIconEx Lib "user32" _
    (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long

Private Declare Function ExtractIconEx Lib "shell32.dll" _
    Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, _
    phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Private Sub cmdBack_Click()
'
' Get the previous icon.
'
If lIndex > 0 Then
    lIndex = lIndex - 1
    Call pGetIcon
End If
End Sub

Private Sub cmdBrowse_Click()
Dim btn    As Button
Dim imgObj As ListImage
'
' Initialize labels. Clear the picture boxes.
'
lIcons = 0
lIndex = 0
lblIcons = 0
lblIcon = 0
lblFile = ""
picSmall.Picture = LoadPicture("")
picLarge.Picture = LoadPicture("")
'
' Remove all toolbar buttons and the
' unbind the ImageList controls.
'
tbLarge.Buttons.Clear
tbLarge.ImageList = Nothing
tbSmall.Buttons.Clear
tbSmall.ImageList = Nothing
'
' Remove all images from the ImageList controls
' and set their size properties.
'
With imgLarge
    .ListImages.Clear
    .ImageHeight = LARGE_ICON
    .ImageWidth = LARGE_ICON
End With

With imgSmall
    .ListImages.Clear
    .ImageHeight = SMALL_ICON
    .ImageWidth = SMALL_ICON
End With
'
' Display the File Open dialog.
' Filter out all files except exe's and dll's.
'
cdlOpen.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
cdlOpen.FileName = ""
cdlOpen.Filter = "Executable Files (*.exe) | *.exe|Application Extension (*.dll) | *.dll"
On Error GoTo CancelButton
cdlOpen.Action = 1
sExeName = cdlOpen.FileName
lblFile = sExeName
'
' Get the total number of Icons in the file.
'
lIcons = ExtractIconEx(sExeName, -1, 0, 0, 0)
'
' Enable various controls.
'
lblIcons = lIcons
cmdBack.Enabled = (lIcons > 1)
cmdNext.Enabled = (lIcons > 1)
lblIcons.Enabled = True
lblIcon.Enabled = True
picSmall.Enabled = True
picLarge.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Frame2.Enabled = True
'
' Dimension the arrays to the number of icons.
' Get the icons' handles.
'
ReDim glLargeIcons(lIcons)
ReDim glSmallIcons(lIcons)
Call pGetIcon
'
' Add the Large icon to the Large ImageList control.
' Bind the large ImageList to the large ToolBar.
' Add a button to the toolbar and populate its ToolTip text.
'
' Note: The "Key" fields of both the ImageList and ToolBar
'       control are set to the same value.  This is what
'       binds a particular image in the ImageList to a
'       given button on the ToolBar control.
'
'           Syntax is:    ...Add(Index, Key, Image)
Set imgObj = imgLarge.ListImages.Add(1, sExeName, picLarge.Image)

With tbLarge
    .ImageList = imgLarge
    ' Syntax is:    ...Add(Index, Key, Caption, Style, Image)
    Set btn = .Buttons.Add(.Buttons.Count + 1, sExeName, , , sExeName)
    .Buttons(1).ToolTipText = sExeName
End With
'
' Repeat for the small icon.
'
Set imgObj = imgSmall.ListImages.Add(1, sExeName, picSmall.Image)
With tbSmall
    .ImageList = imgSmall
    Set btn = .Buttons.Add(.Buttons.Count + 1, sExeName, , , sExeName)
    .Buttons(1).ToolTipText = sExeName
End With

CancelButton:
    'We end up here when hitting Cancel on the Open File dialog.
End Sub

Private Sub cmdNext_Click()
'
' Get the next icon.
'
If lIndex < lIcons - 1 Then
    lIndex = lIndex + 1
    Call pGetIcon
End If
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub Form_Load()
'
' Disable various controls until a file is selected.
'
lIndex = 0

cmdBack.Enabled = False
cmdNext.Enabled = False
lblIcons.Enabled = False
lblIcon.Enabled = False
picSmall.Enabled = False
picLarge.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Frame2.Enabled = False
'
' Align the toolbars to the top of the form.
'
With tbLarge
    .Align = vbAlignTop
    .AllowCustomize = False
    .Wrappable = False
    .BorderStyle = ccNone
End With

With tbSmall
    .Align = vbAlignTop
    .AllowCustomize = False
    .Wrappable = False
    .BorderStyle = ccNone
End With
'
' Set the dimensions of the PictureBox controls where the
' icons will be drawn.  We will use 32x32 and 16x16 icons.
' Each size uses its own PictureBox.
'
picLarge.Height = LARGE_ICON * Screen.TwipsPerPixelY
picLarge.Width = LARGE_ICON * Screen.TwipsPerPixelX
picSmall.Height = SMALL_ICON * Screen.TwipsPerPixelY
picSmall.Width = SMALL_ICON * Screen.TwipsPerPixelX
End Sub



Public Sub pGetIcon()
Dim l As Long
'
' Get the handle of the icon indicated by lIndex.
'
Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)
'
' Draw the icon to respective picturebox control.
'
With picLarge
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)
    .Refresh
End With

With picSmall
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glSmallIcons(lIndex), SMALL_ICON, SMALL_ICON, 0, 0, DI_NORMAL)
    .Refresh
End With
lblIcon = lIndex
End Sub

