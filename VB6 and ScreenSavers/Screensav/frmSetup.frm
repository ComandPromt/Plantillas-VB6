VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options for TheScarms Screen Saver"
   ClientHeight    =   3750
   ClientLeft      =   4935
   ClientTop       =   2655
   ClientWidth     =   5475
   Icon            =   "frmSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frImage 
      Caption         =   "Image"
      Height          =   645
      Left            =   105
      TabIndex        =   28
      Top             =   3045
      Width           =   3375
      Begin VB.OptionButton optImage 
         Caption         =   "Vertex"
         Height          =   225
         Index           =   2
         Left            =   2415
         TabIndex        =   31
         Top             =   315
         Width           =   855
      End
      Begin VB.OptionButton optImage 
         Caption         =   "TheScarms"
         Height          =   225
         Index           =   1
         Left            =   1155
         TabIndex        =   30
         Top             =   315
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optImage 
         Caption         =   "Jordan"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   29
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   435
      Left            =   4515
      TabIndex        =   9
      Top             =   270
      Width           =   855
   End
   Begin VB.Frame frSpeed 
      Caption         =   "Sprite Speed"
      Height          =   645
      Left            =   105
      TabIndex        =   21
      Top             =   2310
      Width           =   5280
      Begin VB.CheckBox chkSpeedRND 
         Caption         =   "Randomize"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
      Begin ComctlLib.Slider sldSpeed 
         Height          =   330
         Left            =   1890
         TabIndex        =   8
         Top             =   210
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         _Version        =   327682
         TickStyle       =   3
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4815
         TabIndex        =   26
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Slow"
         Height          =   195
         Index           =   6
         Left            =   1500
         TabIndex        =   23
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fast"
         Height          =   195
         Index           =   5
         Left            =   4410
         TabIndex        =   22
         Top             =   270
         Width           =   300
      End
   End
   Begin VB.Frame frSpriteSize 
      Caption         =   "Sprite Size %"
      Height          =   645
      Left            =   105
      TabIndex        =   18
      Top             =   1575
      Width           =   5280
      Begin VB.CheckBox chkSizeRND 
         Caption         =   "Randomize"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   270
         Width           =   1125
      End
      Begin ComctlLib.Slider sldSize 
         Height          =   330
         Left            =   1890
         TabIndex        =   6
         Top             =   210
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         _Version        =   327682
         TickStyle       =   3
      End
      Begin VB.Label lblSize 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4815
         TabIndex        =   25
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Small"
         Height          =   195
         Index           =   4
         Left            =   1470
         TabIndex        =   20
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Large"
         Height          =   195
         Index           =   3
         Left            =   4335
         TabIndex        =   19
         Top             =   270
         Width           =   405
      End
   End
   Begin VB.Frame frRefreshRate 
      Caption         =   "Sprite Animation Rate"
      Height          =   645
      Left            =   105
      TabIndex        =   14
      Top             =   840
      Width           =   5280
      Begin VB.CheckBox chkRefreshRND 
         Caption         =   "Randomize"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   315
         Width           =   1125
      End
      Begin ComctlLib.Slider sldRefreshRate 
         Height          =   330
         Left            =   1890
         TabIndex        =   4
         Top             =   210
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         _Version        =   327682
         TickStyle       =   3
      End
      Begin VB.Label lblRefresh 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4815
         TabIndex        =   24
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fast"
         Height          =   195
         Index           =   2
         Left            =   4440
         TabIndex        =   16
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Slow"
         Height          =   195
         Index           =   1
         Left            =   1500
         TabIndex        =   15
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.Frame fSettings 
      Caption         =   "Sprites"
      Height          =   645
      Left            =   105
      TabIndex        =   13
      Top             =   105
      Width           =   4320
      Begin VB.CheckBox chkClearScreen 
         Caption         =   "Clear Screen"
         Height          =   195
         Left            =   2955
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
      Begin VB.PictureBox picCount 
         BackColor       =   &H80000005&
         Height          =   315
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   525
         TabIndex        =   12
         Top             =   210
         Width           =   585
         Begin VB.TextBox txtSprites 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "10"
            Top             =   30
            Width           =   195
         End
         Begin ComCtl2.UpDown udCount 
            Height          =   255
            Left            =   330
            TabIndex        =   27
            Top             =   0
            Width           =   195
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtSprites"
            BuddyDispid     =   196624
            OrigLeft        =   345
            OrigRight       =   540
            OrigBottom      =   255
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
      End
      Begin VB.CheckBox chkTracers 
         Caption         =   "Show Tracers"
         Height          =   195
         Left            =   1425
         TabIndex        =   1
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Count:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4515
      TabIndex        =   11
      Top             =   3200
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   3570
      TabIndex        =   10
      Top             =   3210
      Width           =   855
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkClearScreen_Click()
'
' Save the clear screen option.
'
gbClearScreen = (chkClearScreen.Value = vbChecked)
End Sub

Private Sub chkRefreshRND_Click()
'
' Save the random refresh rate.
'
gbRefreshRND = (chkRefreshRND.Value = vbChecked)
sldRefreshRate.Enabled = Not gbRefreshRND
End Sub
Private Sub chkSizeRND_Click()
'
' Save the random sprite size.
'
gbSizeRND = (chkSizeRND.Value = vbChecked)
sldSize.Enabled = Not gbSizeRND
End Sub
Private Sub chkSpeedRND_Click()
'
' Save the random animation rate.
'
gbSpeedRND = (chkSpeedRND.Value = vbChecked)
sldSpeed.Enabled = Not gbSpeedRND
End Sub
Private Sub chkTracers_Click()
'
' Save the use tracers option.
'
gbUseTracers = (chkTracers.Value = vbChecked)
End Sub

Private Sub cmdAbout_Click()
Dim sTitle As String
Dim sText  As String
'
' Show a Help About dialog.
'
sTitle = "TheScarms Visual Basic/Win32 Code Library#     >>>> www.TheScarms.com <<<<"
sText = vbCrLf & ">>>>   www.TheScarms.com   <<<<" & vbCrLf & vbCrLf
Call ShellAbout(Me.hwnd, sTitle, sText, Me.Icon.Handle)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
'
' Save the current screen saver settings.
'
Call pSaveSettings
Unload Me
End Sub

Private Sub Form_Load()
'
' Load the current screen saver registry settings.
'
Call pLoadSettings
'
' Get the image to display
'
Select Case gsSpriteImage
    Case cIMAGE0
        optImage(0) = True
    Case cIMAGE1
        optImage(1) = True
    Case Else
        optImage(2) = True
End Select
'
' Get the Sprite Count Value.
'
With udCount
    .Max = cMAX_SPRITECOUNT
    .Min = cMIN_SPRITECOUNT
    .Value = glSpriteCount
End With
'
' Get the Refresh Rate Value.
'
With sldRefreshRate
    .Max = cMAX_REFRESHRATE
    .Min = cMIN_REFRESHRATE
    .Value = glRefreshRate
    lblRefresh.Caption = CStr(glRefreshRate)
End With
'
' Get the Sprite Size Value.
'
With sldSize
    .Max = cMAX_SPRITESIZE
    .Min = cMIN_SPRITESIZE
    .Value = glSpriteSize
    lblSize.Caption = CStr(glSpriteSize)
End With
'
' Get the Sprite Speed Value.
'
With sldSpeed
    .Max = cMAX_SPRITESPEED
    .Min = cMIN_SPRITESPEED
    .Value = glSpriteSpeed
    lblSpeed.Caption = CStr(glSpriteSpeed)
End With
'
' Get the Clear Screen Value.
'
If gbClearScreen Then chkClearScreen.Value = vbChecked
'
' Get the Use Tracers Value.
'
If gbUseTracers Then chkTracers.Value = vbChecked
'
' Get the Random Rate Value.
'
If gbRefreshRND Then chkRefreshRND.Value = vbChecked

' Get Random Size Value.
If gbSizeRND Then chkSizeRND.Value = vbChecked
'
' Get the Random Speed Value.
'
If gbSpeedRND Then chkSpeedRND.Value = vbChecked
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSetup = Nothing
End Sub

Private Sub optImage_Click(Index As Integer)
'
' Get the image to display
'
Select Case Index
    Case 0
        gsSpriteImage = cIMAGE0
    Case 2
        gsSpriteImage = cIMAGE2
    Case Else
        gsSpriteImage = cIMAGE1
End Select
End Sub

Private Sub sldRefreshRate_Change()
'
' Save the animation refresh rate.
'
glRefreshRate = sldRefreshRate.Value
lblRefresh.Caption = CStr(glRefreshRate)
End Sub
Private Sub sldRefreshRate_Scroll()
'
' Save the animation refresh rate.
'
glRefreshRate = sldRefreshRate.Value
lblRefresh.Caption = CStr(glRefreshRate)
End Sub
Private Sub sldSize_Change()
'
' Save the active sprite size.
'
glSpriteSize = sldSize.Value
lblSize.Caption = CStr(glSpriteSize)
End Sub
Private Sub sldSize_Scroll()
'
' Save the active sprite size.
'
glSpriteSize = sldSize.Value
lblSize.Caption = CStr(glSpriteSize)
End Sub
Private Sub sldSpeed_Change()
'
' Save the active sprite speed.
'
glSpriteSpeed = sldSpeed.Value
lblSpeed.Caption = CStr(glSpriteSpeed)
End Sub
Private Sub sldSpeed_Scroll()
'
' Save the active sprite speed.
'
glSpriteSpeed = sldSpeed.Value
lblSpeed.Caption = CStr(glSpriteSpeed)
End Sub
Private Sub txtSprites_Change()
'
' Save the active sprite count.
'
glSpriteCount = Val(txtSprites.Text)
End Sub

