VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Class Controls/Owner Drawn Controls"
   ClientHeight    =   4440
   ClientLeft      =   3855
   ClientTop       =   2085
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5775
   Begin VB.Frame Frame4 
      Caption         =   "Owner Drawn Controls"
      Height          =   4215
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   3015
      Begin VB.Frame Frame1 
         Caption         =   "TabStrip With Pictures"
         Height          =   1095
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2775
         Begin ComctlLib.TabStrip TabStrip1 
            Height          =   555
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   979
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   2
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Tab 1"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Tab 2"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Variable Height ListBox"
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   2775
         Begin VB.ListBox lstVariable 
            Height          =   1035
            ItemData        =   "frmSetStyles.frx":0000
            Left            =   120
            List            =   "frmSetStyles.frx":0007
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ComboBox With Pictures"
         Height          =   855
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2775
         Begin VB.ComboBox cboPictures 
            Height          =   315
            ItemData        =   "frmSetStyles.frx":0018
            Left            =   120
            List            =   "frmSetStyles.frx":001A
            TabIndex        =   5
            Text            =   "cboPictures"
            Top             =   360
            Width           =   2535
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Modified Style Controls"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.Frame Frame3 
         Caption         =   "Smooth Progress Bar"
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   2175
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   327682
            Appearance      =   1
            Max             =   10
         End
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   120
            Top             =   1080
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Non-Auto State CheckBox"
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2175
         Begin VB.CheckBox chkNonAuto 
            Caption         =   "Try to check me!"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ComboBox With Tab Stops"
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   2175
         Begin VB.ComboBox cboTabs 
            Height          =   315
            ItemData        =   "frmSetStyles.frx":001C
            Left            =   120
            List            =   "frmSetStyles.frx":0029
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   1875
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

Private Sub chkNonAuto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then MsgBox "Clicked!", vbInformation, "Ha Ha"
End Sub
Private Sub chkNonAuto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then Call chkNonAuto_KeyPress(vbKeySpace)
End Sub

Private Sub Form_Load()
Dim k As Long
'
' Fill the combo to display pictures with
' values. The itemdata value is index in
' the resource file of the picture to display.
'
With cboPictures
    .AddItem "OBM_UPARROW"
    .itemData(.NewIndex) = OBM_UPARROW
    .AddItem "OBM_DNARROW"
    .itemData(.NewIndex) = OBM_DNARROW
    .AddItem "OBM_RGARROW"
    .itemData(.NewIndex) = OBM_RGARROW
    .AddItem "OBM_LFARROW"
    .itemData(.NewIndex) = OBM_LFARROW
    .AddItem "OBM_REDUCE"
    .itemData(.NewIndex) = OBM_REDUCE
    .AddItem "OBM_ZOOM"
    .itemData(.NewIndex) = OBM_ZOOM
    .AddItem "OBM_RESTORE"
    .itemData(.NewIndex) = OBM_RESTORE
    .AddItem "OBM_REDUCED"
    .itemData(.NewIndex) = OBM_REDUCED
    .AddItem "OBM_ZOOMD"
    .itemData(.NewIndex) = OBM_ZOOMD
    .AddItem "OBM_RESTORED"
    .itemData(.NewIndex) = OBM_RESTORED
    .AddItem "OBM_UPARROWD"
    .itemData(.NewIndex) = OBM_UPARROWD
    .AddItem "OBM_DNARROWD"
    .itemData(.NewIndex) = OBM_DNARROWD
    .AddItem "OBM_RGARROWD"
    .itemData(.NewIndex) = OBM_RGARROWD
    .AddItem "OBM_LFARROWD"
    .itemData(.NewIndex) = OBM_LFARROWD
    .AddItem "OBM_MNARROW"
    .itemData(.NewIndex) = OBM_MNARROW
    .AddItem "OBM_COMBO"
    .itemData(.NewIndex) = OBM_COMBO
    .AddItem "OBM_UPARROWI"
    .itemData(.NewIndex) = OBM_UPARROWI
    .AddItem "OBM_DNARROWI"
    .itemData(.NewIndex) = OBM_DNARROWI
    .AddItem "OBM_RGARROWI"
    .itemData(.NewIndex) = OBM_RGARROWI
    .AddItem "OBM_LFARROWI"
    .itemData(.NewIndex) = OBM_LFARROWI
End With
'
' Fill a listbox with font names.
'
For k = 1 To Screen.FontCount
    lstVariable.AddItem Screen.Fonts(k - 1)
Next
'
' Subclass the Parent of the ComboBox, ListBox and TabStrip
' which will be modified as owner drawn controls.  Save the
' address of the original window procedure in the registry.
'
mlWndProc = SetWindowLong(GetParent(cboPictures.hwnd), GWL_WNDPROC, AddressOf fAppWndProc)
Call SaveSetting("OwnerDraw", CStr(GetParent(cboPictures.hwnd)), "WndProcs", CStr(mlWndProc))
'
' Subclass the listbox containing the list of fonts.
'
mlWndProc = SetWindowLong(GetParent(lstVariable.hwnd), GWL_WNDPROC, AddressOf fAppWndProc)
Call SaveSetting("OwnerDraw", CStr(GetParent(lstVariable.hwnd)), "WndProcs", CStr(mlWndProc))
'
' Sub class the tabstrip control.
'
mlWndProc = SetWindowLong(GetParent(TabStrip1.hwnd), GWL_WNDPROC, AddressOf fAppWndProc)
Call SaveSetting("OwnerDraw", CStr(GetParent(TabStrip1.hwnd)), "WndProcs", CStr(mlWndProc))
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
End Sub

Private Sub Timer1_Timer()
'
' Change the value of the progressbar so
' we can watch it scroll smoothly.
'
With ProgressBar1
    If .Value = .Max Then
        .Value = .Min
    Else
        .Value = .Value + 1
    End If
End With
End Sub
