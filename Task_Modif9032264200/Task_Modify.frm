VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task List Modifier 1.04"
   ClientHeight    =   6255
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9690
   Icon            =   "Task_Modify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePP 
      Caption         =   "Properties"
      Height          =   1935
      Left            =   3000
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox PPText 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton PPCmd 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox PPList 
         Height          =   1035
         Left            =   1680
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label PPLabel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame FrameMnu 
      Caption         =   "Menus"
      Height          =   1935
      Left            =   5400
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ListBox MnuList 
         Height          =   285
         Left            =   840
         Style           =   1  'Checkbox
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton MnuCmd 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin MSComctlLib.TreeView MnuTree 
         Height          =   855
         Left            =   1440
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   6
         SingleSel       =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrameCP 
      Caption         =   "Control Properties"
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
      Begin VB.ComboBox CPCombo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CPCmd 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox CPCheck 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox CPText 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label CPLabel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame FrameLst 
      Caption         =   "List Menu"
      Height          =   1815
      Left            =   3000
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox LstText 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton LstButton 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.ListBox LstList 
         Height          =   255
         Left            =   360
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LstLabel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1680
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame FrameMWS 
      Caption         =   "Window Style Menu"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton WSButton 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame FrameWS 
         Caption         =   "Window Styles"
         Height          =   735
         Left            =   2640
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
         Begin VB.ListBox WSList 
            Height          =   285
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame FrameWSX 
         Caption         =   "Extended Window Styles"
         Height          =   615
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         Begin VB.ListBox WSXList 
            Height          =   285
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin MSComctlLib.TreeView TaskTree 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   10292
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   1094
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnuOpt 
      Caption         =   "App Options"
      Begin VB.Menu MnuAOT 
         Caption         =   "Always On Top"
      End
      Begin VB.Menu MnuLegend 
         Caption         =   "Show Legend"
      End
      Begin VB.Menu mnuseparator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoreOptions 
         Caption         =   "More Options..."
      End
   End
   Begin VB.Menu TaskMenu_Items 
      Caption         =   "Selected Item Options"
      Begin VB.Menu TreeMenuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuSearchFor 
         Caption         =   "Search For"
         Begin VB.Menu mnuFindByExec 
            Caption         =   "Executable"
         End
         Begin VB.Menu mnuFindByText 
            Caption         =   "Text"
         End
         Begin VB.Menu mnuFindByHandle 
            Caption         =   "Handle"
         End
         Begin VB.Menu mnuFindByFindAgain 
            Caption         =   "Find Again"
            Shortcut        =   {F3}
            Visible         =   0   'False
         End
      End
      Begin VB.Menu TreeMenuShowWindow 
         Caption         =   "&Show Window"
         Begin VB.Menu TreeMenuCProps 
            Caption         =   "&Control Attributes"
         End
         Begin VB.Menu TreeMenuShowMenus 
            Caption         =   "&Menus"
         End
         Begin VB.Menu TreeMenuWindowStyles 
            Caption         =   "&Window Styles"
         End
         Begin VB.Menu TreeMenuListItems 
            Caption         =   "&ListItems"
         End
         Begin VB.Menu TreeMenuProperties 
            Caption         =   "&Property Bag"
         End
      End
      Begin VB.Menu TreeMenuActivate 
         Caption         =   "&Activate"
      End
      Begin VB.Menu mnuseparator1 
         Caption         =   "-"
      End
      Begin VB.Menu TreeMenuEndTask 
         Caption         =   "&Kill Control"
      End
      Begin VB.Menu mnuseparator2 
         Caption         =   "-"
      End
      Begin VB.Menu TreeMenuLaunch 
         Caption         =   "Launch External App"
      End
   End
   Begin VB.Menu mnuFindObj 
      Caption         =   "Find object with mouse"
   End
   Begin VB.Menu mnuSecretOption 
      Caption         =   "Secret Option"
      Visible         =   0   'False
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Programmed in Win2k. i hope everything is compatable with 9x.

'Set this to False if you would like to Enable Caption changing and being able to resize
'the form with api.
Option Explicit
#Const Stop_Changes = True

'Some of the API procedures defined in this app are unused...

'Right Click on Treeview Items to show the options and commands

'Note: maybe its just a Win2k problem only but on this app do not confuse Desktop(in green) with Progman(The actual desktop window that has icons)
'       I do not know why the are different, but tampering with the desktop(in green)(GetDesktopWindow API) can actually
'       disable the user from doing anything until the computer is restarted.(Progman is started with explorer)


'AuthorMode Is needed to be able to access Write features for this app.

Private Const Evnt_FindNullHwnd  As Long = 1
'constants for CommandButton Width/Height
Private Const bWidth As Long = 1335
Private Const bHeight As Long = 375
Private Const bLeft As Long = 360
Private CPFlag As Boolean 'Used to stop the code from indirectly changing
                          'the index of the combobox and causing events..

Public Property Let AlwaysOnTop(Flg As Boolean)

  Dim HWND_MyFlag As Long

    If Flg Then
        HWND_MyFlag = HWND_TOPMOST
      Else 'NOT FLG...
        HWND_MyFlag = HWND_NOTOPMOST
    End If
    SetWindowPos Me.hwnd, HWND_MyFlag, 0, 0, 0, 0, AOT_Flags
    MnuAOT.Checked = Flg

End Property

Private Sub CPCheck_Click(Index As Integer)

    If AuthorMode Then
        Select Case Index
          Case 0 'Visible
            If CPFlag Then
                '2 min ---- 3 max ---- 4 normal
                ShowWindow SelectedNodeHwnd, CPCheck(0).Value * (GetWindowState(SelectedNodeHwnd) + 2)
                Me.ZOrder
            End If
          Case 1 'Enabled
            SetControlEnabled SelectedNodeHwnd, Abs(CPCheck(1).Value)
        End Select
    End If

End Sub

Private Sub CPCmd_Click(Index As Integer)
'TaskMenuId Windows
'3 Control Propertys

  Dim tmp As Integer

    Select Case Index
      Case 0 'Close Menu
        TaskMenuID = 0
        FrameCP.Visible = False
        TaskTree.Enabled = True

        For tmp = 1 To 17
            Unload CPText(tmp)
            Unload CPLabel(tmp)
        Next tmp
        Unload CPLabel(18)
        Unload CPLabel(19)
        Unload CPCombo(1)
        Unload CPCheck(1)
    End Select

End Sub

Private Sub CPCombo_Click(Index As Integer)

  Dim ProcID As Long

    If AuthorMode Then
        Select Case Index
          Case 0 'WindowState
            If CPFlag Then
                'Set and retrieve the new windowstate.
                CPCombo(0).ListIndex = SetWindowState(SelectedNodeHwnd, CPCombo(0).ListIndex + 2)
                Me.ZOrder 'call ontop for this apps visual safety
            End If
          Case 1 'Process Priority

            Get_Thread_ProcessID SelectedNodeHwnd, ProcID
            CPCombo(1).ListIndex = SetProcessPriority(ProcID, CPCombo(1).ListIndex)
        End Select
    End If

End Sub

Private Sub CPText_GotFocus(Index As Integer)

    If AuthorMode Then
        Select Case Index
          Case 1, 2, 8, 9, 10, 11, 12
            SetPropsText CPText(Index), True, vbBlue
        End Select
    End If

End Sub

Private Sub CPText_KeyPress(Index As Integer, KeyAscii As Integer)

    If AuthorMode Then
  Dim rLeft As Long, rTop As Long, rWidth As Long, rHeight As Long

        If KeyAscii = 13 Then
            Select Case Index
              Case 1 'Caption
                CPText(1).Text = SetText(SelectedNodeHwnd, CPText(1).Text)
                TaskTree.Nodes.item(TaskTree.SelectedItem.Key).Text = CPText(1).Text
              Case 2 'Parent
                CPText(2).Text = AssignParent(SelectedNodeHwnd, CLng(Val(CPText(Index).Text)))
              Case 8 'PWchar
                SetNewPwChar SelectedNodeHwnd, CByte(Val(CPText(8).Text))
              Case 9, 10, 11, 12 '(Top,Left,Width,Height)
                MoveControl SelectedNodeHwnd, CLng(Val(CPText(9).Text)), CLng(Val(CPText(10).Text)), CLng(Val(CPText(11).Text)), CLng(Val(CPText(12).Text))
                GetControlRect SelectedNodeHwnd, rTop, rLeft, rWidth, rHeight
                CPText(9) = rTop
                CPText(10) = rLeft
                CPText(11) = rWidth
                CPText(12) = rHeight
            End Select
        End If
    End If

End Sub

Private Sub CPText_LostFocus(Index As Integer)

    If AuthorMode Then
        Select Case Index
          Case 1, 2, 8, 9, 10, 11, 12
            SetPropsText CPText(Index), True, vbBlack
        End Select
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 And FindAgain Then
        mnuFindByFindAgain_Click
    End If
    If KeyCode = vbKeyShift Then
        mnuSecretOption.Visible = True
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        mnuSecretOption.Visible = False
    End If
End Sub

Private Sub Form_Load()

  Dim X As Integer
  Dim LaunchFLen As Integer
  Dim fdir As String
  Dim buff As String
    
'use for compilation - if you'd like to change this look at the constant at the top of this code
#If Stop_Changes = True Then
    NoSizing (Me.hwnd)
    AllowCaptionChange Me.hwnd, False
#End If
    
    RefreshTask
    TaskTree_NodeClick TaskTree.Nodes.item(1)
    'lets open our .dat file now to get our saved variables
    fdir = Dir$(App.Path & "\taskmod.dat")
    If fdir <> "" Then
        Open fdir For Binary As #1
        Get #1, 1, AuthorMode
        Get #1, , Showicons
        Get #1, , SearchForWindows
        Get #1, , LaunchPar
        Get #1, , LaunchFLen
        buff = String$(LaunchFLen, 0)
        Get #1, , buff
        LaunchFile = buff
        Close #1
    End If
    If SearchForWindows Then
        SetTimer Me.hwnd, Evnt_FindNullHwnd, 1000, AddressOf TimerProc
    End If
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    End If
    If Me.WindowState = vbNormal Then
        Me.Width = 9810
        Me.Height = 6675
        TaskTree.Width = 9710
        TaskTree.Height = 6025
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

  'TaskMenuID is used becuz of users trying to click the form_close button
  'instead of the Close button provided for them. therefore the form will not
  'close if any options are open.
'TaskMenuId Windows
'1 Window Styles
'2 List Box Propertys
'3 Control Propertys
'4 Menu's
'5 Property Bag
'6 Options

'NOTE TO AUTHOR: if adding/removing buttons; remember that the indexes will change
'                on this function below when trying to close, and in the close subs.
Cancel = True 'stop the Form unload for now
'basically if an option is open, close it, if the main window is showing then end app
Select Case TaskMenuID
    Case 0 'Ok to close form
        'clean up before unload
        LockWindowUpdate Me.hwnd
        AllowCaptionChange Me.hwnd, True
        Unload Me 'unload the main form
        TaskTree.Nodes.Clear 'clear the nodes
        Set TreeX = Nothing ' clear memory
        LockWindowUpdate 0 'unlock view
        Cancel = False 'let the app unload
        End
    Case 1: WSButton_Click 1 ' the 1 is the Index for the close button
    Case 2: LstButton_Click 3 ' the 3 is the index for the close button
    Case 3: CPCmd_Click 0 ' the 0 is the index for the close button
    Case 4: MnuCmd_Click 3 ' the 3 is the index for the close button
    Case 5: PPCmd_Click 3 ' the 3 is the index for the close button
    Case 6: Unload Options 'unload the options form
End Select
End Sub

Private Sub LoadTaskList()

  Dim DeskTophwnd As String
  Dim SWindowText As String
  Dim Nodx As Node
  Dim X As Long
  Dim cur As Long
  Dim tmpcounter As Long
    
    Me.MousePointer = vbHourglass
    tmpcounter = 0
    curhwnd = GetDesktopWindow()
    TreeX.Clear
    DeskTophwnd = CStr(curhwnd)
    SWindowText = Space$(255)
    GetComputerName SWindowText, 255
    SWindowText = Left$(SWindowText, InStr(SWindowText, Chr$(0)) - 1)
    Set Nodx = TaskTree.Nodes.Add(, , "t0", SWindowText)
    Set Nodx = TaskTree.Nodes.Add(, , "t" & DeskTophwnd, "Desktop")
    TaskTree.Nodes.item(1).ForeColor = RGB(0, 175, 0)
    TaskTree.Nodes.item(2).ForeColor = RGB(0, 175, 0)
    GetAllChildren curhwnd
    TreeX.RemoveNode CLng(DeskTophwnd)
    With ImageList1
        .ListImages.Clear
        .ImageHeight = 16
        .ImageWidth = 16
    End With
    For X = 2 To TreeX.GetCount
        cur = CLng(TreeX.GetItem(X))
        Set Nodx = TaskTree.Nodes.Add("t" & CStr(GetParent(cur)), tvwChild, _
            "t" & CStr(cur), GetFriendlyName(cur))
        If Showicons Then
            PicIcon.Cls
            DrawIcon PicIcon.hdc, 0, 0, DetermineBestIcon(cur)
            ImageList1.ListImages.Add , , PicIcon.Image
            Nodx.Image = ImageList1.ListImages.Count
            ''never clear listimages or picture will not appear on treeview
        End If
        If cur = Me.hwnd Then
            Nodx.ForeColor = vbBlue
          ElseIf IsWindowVisible(cur) = 0 Then
            Nodx.ForeColor = RGB(127, 127, 127)
        End If
    Next X
    SortNodes TaskTree
    TaskTree.Nodes(1).Expanded = True
    'clear some mem...
    Set Nodx = Nothing
    Me.MousePointer = vbDefault

End Sub

Private Sub LstButton_Click(Index As Integer)
'TaskMenuId Windows
'2 List Box Propertys

  Dim X As Integer

    Select Case Index
      Case 0 'Refresh
        For X = 1 To 2
            Unload LstButton(X)
            Unload LstText(X)
            Unload LstLabel(X)
        Next X
        Unload LstButton(3)
        TreeMenuListItems_Click
      Case 1 'Add item
        If AuthorMode Then
            If LstText(1) = "" Then
                LstText(1) = "0"
            End If
            If LstText(2) = "" Then
                LstText(2) = " "
            End If
            LstAddItem LstList.hwnd, SelectedNodeHwnd, LstText(2), CLng(LstText(1))
        End If
      Case 2 'Remove item
        If AuthorMode Then
            If LstList.ListIndex > -1 Then
                LstRemoveItem LstList.hwnd, SelectedNodeHwnd, LstList.ListIndex
            End If
        End If
      Case 3 'Close Menu
        TaskMenuID = 0
        FrameLst.Visible = False
        TaskTree.Enabled = True
        For X = 1 To 2
            Unload LstButton(X)
            Unload LstText(X)
            Unload LstLabel(X)
        Next X
        Unload LstButton(3)
    End Select
    LstText(0).Text = LstList.ListCount

End Sub

Private Sub LstList_Click()

    LstText(0).Text = LstList.ListCount
    LstText(2).Text = LstList.List(LstList.ListIndex)
    LstText(1).Text = LstList.ItemData(LstList.ListIndex)

End Sub

Private Sub LstText_GotFocus(Index As Integer)

    If AuthorMode Then
        Select Case Index
          Case 1, 2
            SetPropsText LstText(Index), True, vbBlue
        End Select
    End If

End Sub

Private Sub LstText_KeyPress(Index As Integer, KeyAscii As Integer)

  Dim LI As Long

    If AuthorMode Then
        If KeyAscii = 13 Then

            LI = LstList.ListIndex
            If LI > -1 Then
                LstReplaceItem LstList.hwnd, SelectedNodeHwnd, LI, LstText(2).Text, Val(LstText(1).Text)
            End If
        End If
    End If

End Sub

Private Sub LstText_LostFocus(Index As Integer)

    If AuthorMode Then
        Select Case Index
          Case 1, 2
            SetPropsText LstText(Index), True, vbBlack
        End Select
    End If

End Sub

Private Sub MnuAbout_Click()

    MsgBox "Programmed By Billy Conner 2001", vbOKOnly, "About"

End Sub

Private Sub MnuAOT_Click()

    If MnuAOT.Checked = 0 Then
        Me.AlwaysOnTop = True
      Else
        Me.AlwaysOnTop = False
    End If

End Sub

Private Sub MnuCmd_Click(Index As Integer)
'TaskMenuId Windows
'4 Menu's

  Dim dKey As String
  Dim tmp As Integer
  Dim st As Variant

    dKey = MnuTree.SelectedItem.Key
    st = Split(dKey, ":")
    Select Case Index
      Case 0 'refresh
        TaskTree.Enabled = True
        For tmp = 1 To 3
            Unload MnuCmd(tmp)
        Next tmp
        TreeMenuShowMenus_Click
      Case 1 'removeitem
        If AuthorMode Then
            If Len(st(0)) = 1 Then
                RemoveMenuItem SelectedNodeHwnd, st(0) = "S", CLng(st(1))
                MnuTree.Nodes.Remove (MnuTree.SelectedItem.Key)
            End If
        End If
      Case 2 'Run Item
        If AuthorMode Then
            If UBound(st) = 3 Then
                RunMenuItem SelectedNodeHwnd, CLng(st(1))
            End If
        End If
      Case 3 'Close Menu
        TaskMenuID = 0
        FrameMnu.Visible = False
        TaskTree.Enabled = True
        For tmp = 1 To 3
            Unload MnuCmd(tmp)
        Next tmp
    End Select

End Sub

Private Sub mnuFindByExec_Click()

    mnuFindByFindAgain.Visible = FindText(First, TaskTree, FINDBY_EXECUTABLE)

End Sub

Private Sub mnuFindByFindAgain_Click()

    mnuFindByFindAgain.Visible = FindText(FindNext, TaskTree)

End Sub

Private Sub mnuFindByHandle_Click()

    mnuFindByFindAgain.Visible = FindText(First, TaskTree, FINDBY_HANDLE)

End Sub

Private Sub mnuFindByText_Click()

    mnuFindByFindAgain.Visible = FindText(First, TaskTree, FINDBY_TEXT)

End Sub

Private Sub mnuFindObj_Click()
SetTimer Me.hwnd, Evnt_Countdown, 1000, AddressOf Proc_FHFMO_CountDown
SetTimer Me.hwnd, Evnt_FindHwndFromMouseOver, 100, AddressOf Proc_FindHwndFromMouseOver
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub mnuSecretOption_Click()
MsgBox ":)"
End Sub

Private Sub MnuVisibleOnly_Click()

End Sub

Private Sub TreeMenuLaunch_Click()

   Dim FilePar As String
   Dim Thread As Long, ProcID As Long
   
    If Dir(LaunchFile) <> "" Then
        Select Case LaunchPar
            Case 0 'none
               'do nothing
            Case 1 'Handle
                FilePar = CStr(SelectedNodeHwnd)
            Case 2 'Parent Handle
                FilePar = CStr(GetParent(SelectedNodeHwnd))
            Case 3 'Process ID
                Thread = Get_Thread_ProcessID(SelectedNodeHwnd, ProcID)
                FilePar = CStr(ProcID)
            Case 4 'Exec File
                FilePar = GetExeFromHandle(SelectedNodeHwnd)
        End Select
    ShellExecute Me.hwnd, "open", LaunchFile, FilePar, App.Path, vbNormalFocus
    End If
End Sub

Private Sub MnuLegend_Click()

  'Load Legend

    AssignParent Legend.hwnd, TaskTree.hwnd
    Legend.Move TaskTree.Width - Legend.Width - 300, TaskTree.Top
    Legend.Show

End Sub

Private Sub mnuMoreOptions_Click()
    TaskMenuID = 6
    Load Options
    AssignParent Options.hwnd, Me.hwnd
    Options.Move TaskTree.Left, TaskTree.Top, TaskTree.Width, TaskTree.Height
    Options.Show
End Sub

Private Sub MnuTree_NodeClick(ByVal Node As MSComctlLib.Node)

  Dim ms As Long
  Dim dKey As String
  Dim st As Variant
  
    dKey = MnuTree.SelectedItem.Key
    st = Split(dKey, ":")
    If Left$(st(0), 1) = "S" Then
        MnuCmd(2).Enabled = False
      Else
        MnuCmd(2).Enabled = True
    End If
    If UBound(st) = 3 Then
        ms = CLng(st(2))
        CheckMenuStats MnuList, ms
    End If

End Sub

Private Sub PPCmd_Click(Index As Integer)
'TaskMenuId Windows
'5 Property Bag

  Dim tmp As Integer

    Select Case Index
      Case 0 'Refresh
        For tmp = 1 To 2
            Unload PPCmd(tmp)
            Unload PPText(tmp)
            Unload PPLabel(tmp)
        Next tmp
        Unload PPCmd(3)
        TreeMenuProperties_Click
      Case 1 'Add Item
        If AuthorMode Then
            If IsStringNumeric(PPText(2).Text) = False Then
                PPText(2).Text = "0"
            End If
            Add_Prop SelectedNodeHwnd, PPText(1), PPText(2)
        End If
      Case 2 'Remove Item
        If AuthorMode Then
            Delete_Prop SelectedNodeHwnd, PPText(1)
        End If
      Case 3 'Close Menu
        TaskMenuID = 0
        FramePP.Visible = False
        TaskTree.Enabled = True
        For tmp = 1 To 2
            Unload PPCmd(tmp)
            Unload PPText(tmp)
            Unload PPLabel(tmp)
        Next tmp
        Unload PPCmd(3)
    End Select

End Sub

Private Sub PPList_Click()

    PPText(1).Text = PPList.List(PPList.ListIndex)
    PPText(2).Text = Get_Prop(SelectedNodeHwnd, PPText(1).Text)

End Sub

Private Sub PPText_GotFocus(Index As Integer)

    If AuthorMode Then
        Select Case Index
          Case 1, 2
            SetPropsText PPText(Index), True, vbBlue
        End Select
    End If

End Sub

Private Sub PPText_KeyPress(Index As Integer, KeyAscii As Integer)

    If AuthorMode Then
        If Index <> 0 Then
            If KeyAscii = 13 Then
                If IsStringNumeric(PPText(2).Text) = False Then
                    PPText(2).Text = "0"
                End If
                Add_Prop SelectedNodeHwnd, PPText(1), PPText(2)
            End If
        End If
    End If

End Sub

Private Sub PPText_LostFocus(Index As Integer)

    If AuthorMode Then
        Select Case Index
          Case 1, 2
            SetPropsText PPText(Index), True, vbBlack
        End Select
    End If

End Sub

Public Sub RefreshTask()

    TreeMenuRefresh_Click

End Sub

Private Sub SortNodes(DaTree As TreeView)

  Dim n As Node

    For Each n In DaTree.Nodes
        n.Sorted = True
    Next n

End Sub

Private Sub TaskTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu TaskMenu_Items
    End If

End Sub

Private Sub TaskTree_NodeClick(ByVal Node As MSComctlLib.Node)

    SelectedNodeKey = TaskTree.SelectedItem.Key
    SelectedNodeHwnd = Mid$(TaskTree.SelectedItem.Key, 2)
    TreeMenuListItems.Enabled = IsList(SelectedNodeHwnd)
    TreeMenuShowMenus.Enabled = IsMenu(SelectedNodeHwnd)
    TreeMenuListItems.Caption = "&ListItems (" & CStr(GetListItemCount(SelectedNodeHwnd)) & ")"
    TreeMenuProperties.Caption = "&Property Bag (" & CStr(GetPropCount(SelectedNodeHwnd)) & ")"
    If TaskTree.Nodes.item(1).Selected = False And TaskTree.Nodes.item(2).Selected = False Then
        TreeMenuLaunch.Visible = True
        TreeMenuEndTask.Visible = True
        TreeMenuActivate.Visible = True
        TreeMenuShowWindow.Visible = True
      Else
        TreeMenuLaunch.Visible = False
        TreeMenuEndTask.Visible = False
        TreeMenuActivate.Visible = False
        TreeMenuShowWindow.Visible = False
    End If

End Sub

Private Sub TreeMenuActivate_Click()

    BringWindowToTop SelectedNodeHwnd

End Sub

Private Sub TreeMenuCProps_Click()

  Dim SClassName As String * 255
  Dim rLeft As Long, rTop As Long, rWidth As Long, rHeight As Long
  Dim Thread As Long, ProcID As Long
  Dim tmp As Integer

    TaskMenuID = 3
    TaskTree.Enabled = False
    LockWindowUpdate FrameCP.hwnd
    CPFlag = False
    CPLabel(0).FontSize = 10
    CPLabel(0).Alignment = vbRightJustify
    SetProps FrameCP, 0, 0, 9710, TaskTree.Height, True
    GetClassName SelectedNodeHwnd, SClassName, 255
    SetProps CPLabel(0), 360, 120, 1215, 255, True, , "Class:"
    SetProps CPText(0), 360, 1440, 6375, 285, True, True, SClassName
    Load CPLabel(1)
    SetProps CPLabel(1), 720, 240, 1095, 255, True, , "Caption:"
    Load CPText(1)
    SetProps CPText(1), 720, 1440, 6375, 285, True, (AuthorMode = False), GetText(SelectedNodeHwnd)
    Load CPLabel(2)
    SetProps CPLabel(2), 1320, 240, 1095, 255, True, , "Parent:"
    Load CPText(2)
    SetProps CPText(2), 1320, 1440, 1000, 285, True, (AuthorMode = False), GetParent(SelectedNodeHwnd)
    Load CPLabel(3)
    SetProps CPLabel(3), 1680, 120, 1215, 255, True, , "Handle:"
    Load CPText(3)
    SetProps CPText(3), 1680, 1440, 1000, 285, True, True, CStr(SelectedNodeHwnd)
    Load CPLabel(4)
    SetProps CPLabel(4), 2040, 120, 1215, 255, True, , "Instance:"
    Load CPText(4)
    SetProps CPText(4), 2040, 1440, 1000, 285, True, True, GetWndTypeVal(SelectedNodeHwnd, GWW_HINSTANCE)

    Thread = Get_Thread_ProcessID(SelectedNodeHwnd, ProcID)
    Load CPLabel(5)
    SetProps CPLabel(5), 2400, 120, 1215, 255, True, , "Thread ID:"
    Load CPText(5)
    SetProps CPText(5), 2400, 1440, 1000, 285, True, True, CStr(Thread)
    Load CPLabel(6)
    SetProps CPLabel(6), 2760, 120, 1215, 255, True, , "Process ID:"
    Load CPText(6)
    SetProps CPText(6), 2760, 1440, 1000, 285, True, True, CStr(ProcID)
    Load CPLabel(7)
    SetProps CPLabel(7), 3120, 120, 1215, 255, True, , "Window ID:"
    Load CPText(7)
    SetProps CPText(7), 3120, 1440, 1000, 285, True, True, GetWndTypeVal(SelectedNodeHwnd, GWL_ID)
    Load CPLabel(8)
    SetProps CPLabel(8), 3840, 840, 1000, 255, True, , "Pw Char:"
    Load CPText(8)
    SetProps CPText(8), 3840, 1920, 495, 285, True, (AuthorMode = False), Chr$(GetPassWordChar(SelectedNodeHwnd))
    GetControlRect SelectedNodeHwnd, rTop, rLeft, rWidth, rHeight
    Load CPLabel(9)
    SetProps CPLabel(9), 1320, 3120, 735, 255, True, , "Top:"
    Load CPText(9)
    SetProps CPText(9), 1320, 3960, 1000, 285, True, (AuthorMode = False), CStr(rTop)
    Load CPLabel(10)
    SetProps CPLabel(10), 1680, 3120, 735, 255, True, , "Left:"
    Load CPText(10)
    SetProps CPText(10), 1680, 3960, 1000, 285, True, (AuthorMode = False), CStr(rLeft)
    Load CPLabel(11)
    SetProps CPLabel(11), 2040, 3120, 735, 255, True, , "Width:"
    Load CPText(11)
    SetProps CPText(11), 2040, 3960, 1000, 285, True, (AuthorMode = False), CStr(rWidth)
    Load CPLabel(12)
    SetProps CPLabel(12), 2400, 3120, 735, 255, True, , "Height:"
    Load CPText(12)
    SetProps CPText(12), 2400, 3960, 1000, 285, True, (AuthorMode = False), CStr(rHeight)
    Load CPLabel(13)
    SetProps CPLabel(13), 3120, 2880, 1000, 255, True, , "Style:"
    Load CPText(13)
    SetProps CPText(13), 3120, 3960, 1000, 285, True, True, GetWndTypeVal(SelectedNodeHwnd, GWL_STYLE)
    Load CPLabel(14)
    SetProps CPLabel(14), 3480, 2880, 1000, 255, True, , "Ex-Style:"
    Load CPText(14)
    SetProps CPText(14), 3480, 3960, 1000, 285, True, True, GetWndTypeVal(SelectedNodeHwnd, GWL_EXSTYLE)
    Load CPLabel(15)
    SetProps CPLabel(15), 3840, 2880, 1000, 255, True, , "hDC:"
    Load CPText(15)
    SetProps CPText(15), 3840, 3960, 1000, 285, True, True, CStr(GetDC(SelectedNodeHwnd))
    Load CPLabel(16)
    SetProps CPLabel(16), 2760, 5640, 2175, 255, True, , "Menu Handle:"
    Load CPText(16)
    SetProps CPText(16), 2760, 7920, 1000, 285, True, True, GetMenu(SelectedNodeHwnd)
    Load CPLabel(17)
    SetProps CPLabel(17), 3120, 5640, 2175, 255, True, , "System Menu Handle:"
    Load CPText(17)
    SetProps CPText(17), 3120, 7920, 1000, 285, True, True, GetSystemMenu(SelectedNodeHwnd, 0)
    SetProps CPCheck(0), 1320, 6600, 1095, 255, True, (AuthorMode = False), "Visible"
    Load CPCheck(1)
    SetProps CPCheck(1), 1320, 7800, 1095, 255, True, (AuthorMode = False), "Enabled"
    Load CPLabel(18)
    SetProps CPLabel(18), 1680, 5640, 1575, 255, True, , "Window State:"
    Load CPLabel(19)
    SetProps CPLabel(19), 2160, 5640, 1575, 255, True, , "Process Priority:"
    SetProps CPCombo(0), 1680, 7320, 1575, , True, (AuthorMode = False)
    Load CPCombo(1)
    SetProps CPCombo(1), 2160, 7320, 1575, , True, (AuthorMode = False)
    SetProps CPCmd(0), FrameCP.Height - 600, 480, bWidth, bHeight, True, , "Close This Menu"
    For tmp = 0 To 17
        Select Case tmp
          Case 1, 2, 8, 9, 10, 11, 12
            SetPropsText CPText(tmp), (AuthorMode = True)
        End Select
    Next tmp
    CPCheck(0).FontSize = 10
    CPCheck(0).Value = Abs(IsWindowVisible(SelectedNodeHwnd))
    CPCheck(1).Value = Abs(IsWindowEnabled(SelectedNodeHwnd))
    CPCombo(0).Clear
    CPCombo(0).AddItem "Minimized"
    CPCombo(0).AddItem "Maximized"
    CPCombo(0).AddItem "Normal"
    CPCombo(1).Clear
    CPCombo(1).AddItem "Low"
    CPCombo(1).AddItem "Below Normal"
    CPCombo(1).AddItem "Normal"
    CPCombo(1).AddItem "Above Normal"
    CPCombo(1).AddItem "High"
    CPCombo(1).AddItem "Realtime"
    CPCombo(1).ListIndex = ConvertPriorityToComboBoxVal(GetProcessPriority(ProcID))
    CPCombo(0).ListIndex = GetWindowState(SelectedNodeHwnd)
    CPFlag = True
    LockWindowUpdate 0

End Sub

Private Sub TreeMenuEndTask_Click()

    If AuthorMode Then
  Dim IsDead As Integer

        IsDead = KillWindow(SelectedNodeHwnd)
        Select Case IsDead
          Case 0
            MsgBox "All attempts on killing the control failed."
          Case Else
            If IsDead = 1 Then
                MsgBox "The window closed successfully."
              ElseIf IsDead = 2 Then
                MsgBox "The window had to be destroyed."
              ElseIf IsDead = 3 Then
                MsgBox "The window could not be closed so the process was terminated."
            End If
            RefreshTask
        End Select
    End If

End Sub

Private Sub TreeMenuListItems_Click()

  Dim tmp As Integer

    TaskMenuID = 2
    TaskTree.Enabled = False
    LockWindowUpdate FrameLst.hwnd
    SetProps FrameLst, 0, 0, 9710, TaskTree.Height, True
    SetProps LstList, 120, 6120, 3550, FrameLst.Height - 200, True
    LstList.Clear
    CopyListToList SelectedNodeHwnd, LstList.hwnd
    SetProps LstButton(0), 360, bLeft, bWidth, bHeight, True, , "Refresh"
    Load LstButton(1)
    SetProps LstButton(1), 960, bLeft, bWidth, bHeight, True, (AuthorMode = False), "Add Item"
    Load LstButton(2)
    SetProps LstButton(2), 1560, bLeft, bWidth, bHeight, True, (AuthorMode = False), "Remove Item"
    Load LstButton(3)
    SetProps LstButton(3), FrameLst.Height - 600, bLeft, bWidth, bHeight, True, , "Close This Menu"
    SetProps LstText(0), 600, 3000, 2700, 285, True, True, LstList.ListCount
    Load LstText(1)
    SetProps LstText(1), 1400, 3000, 2700, 285, True, (AuthorMode = False), ""
    Load LstText(2)
    SetProps LstText(2), 2200, 3000, 2700, 285, True, (AuthorMode = False), ""
    SetProps LstLabel(0), 360, 3000, 2700, 225, True, , "Number Of Items"
    Load LstLabel(1)
    SetProps LstLabel(1), 1160, 3000, 2700, 225, True, , "Item Data"
    Load LstLabel(2)
    SetProps LstLabel(2), 1960, 3000, 2700, 225, True, , "List Item Text"

    For tmp = 0 To 2
        Select Case tmp
          Case 1, 2
            SetPropsText LstText(tmp), (AuthorMode = True)
        End Select
    Next tmp
    LockWindowUpdate 0

End Sub

Private Sub TreeMenuProperties_Click()

  Dim tmp As Long

    TaskMenuID = 5
    TaskTree.Enabled = False
    LockWindowUpdate FramePP.hwnd
    SetProps FramePP, 0, 0, 9710, TaskTree.Height, True
    SetProps PPList, 120, 6120, 3550, FramePP.Height - 200, True
    SetProps PPCmd(0), 360, bLeft, bWidth, bHeight, True, , "Refresh"
    Load PPCmd(1)
    SetProps PPCmd(1), 960, bLeft, bWidth, bHeight, True, (AuthorMode = False), "Add Item"
    Load PPCmd(2)
    SetProps PPCmd(2), 1560, bLeft, bWidth, bHeight, True, (AuthorMode = False), "Remove Item"
    Load PPCmd(3)
    SetProps PPCmd(3), FramePP.Height - 600, bLeft, bWidth, bHeight, True, , "Close This Menu"
    SetProps PPText(0), 600, 3000, 2700, 285, True, True, ""
    Load PPText(1)
    SetProps PPText(1), 1400, 3000, 2700, 285, True, (AuthorMode = False), ""
    Load PPText(2)
    SetProps PPText(2), 2200, 3000, 2700, 285, True, (AuthorMode = False), ""
    SetProps PPLabel(0), 360, 3000, 2700, 225, True, , "Number Of Properties"
    Load PPLabel(1)
    SetProps PPLabel(1), 1160, 3000, 2700, 225, True, , "Property Name"
    Load PPLabel(2)
    SetProps PPLabel(2), 1960, 3000, 2700, 225, True, , "Property Value"

    PPList.Clear
    GetPropList SelectedNodeHwnd
    If PPList.ListCount Then
        PPList.ListIndex = 0
    End If
    For tmp = 1 To 2
        SetPropsText PPText(tmp), (AuthorMode = True)
    Next tmp
    LockWindowUpdate 0

End Sub

Private Sub TreeMenuRefresh_Click()

  Dim n As Node

    Me.MousePointer = vbHourglass
    If SelectedNodeKey <> "" Then
        SelectedNodeKey = TaskTree.SelectedItem.Key
    End If
    If IsWindow(SelectedNodeHwnd) = 0 Then
        SelectedNodeKey = "t0"
    End If
    LockWindowUpdate TaskTree.hwnd
    TaskTree.Nodes.Clear
    LoadTaskList

    For Each n In TaskTree.Nodes
        If n.Key = SelectedNodeKey Then
            n.Selected = True
            Exit For
        End If
    Next n
    LockWindowUpdate (0)
    Me.MousePointer = vbDefault
    
End Sub

Private Sub TreeMenuShowMenus_Click()

  Dim SMnu As Long, BMnu As Long

    TaskMenuID = 4
    TaskTree.Enabled = False
    LockWindowUpdate FrameMnu.hwnd
    SetProps FrameMnu, 0, 0, 9710, TaskTree.Height, True
    SetProps MnuList, 120, 2000, 2100, FrameMnu.Height - 200, True
    SetProps MnuTree, 120, 4100, 5550, MnuList.Height, True
    SetProps MnuCmd(0), 360, bLeft, bWidth, bHeight, True, , "Refresh"
    Load MnuCmd(1)
    SetProps MnuCmd(1), 960, bLeft, bWidth, bHeight, True, (AuthorMode = False), "Remove Item"
    Load MnuCmd(2)
    SetProps MnuCmd(2), 1560, bLeft, bWidth, bHeight, True, (AuthorMode = False), "Run Item"
    Load MnuCmd(3)
    SetProps MnuCmd(3), FrameMnu.Height - 600, bLeft, bWidth, bHeight, True, , "Close This Menu"
    MnuTree.Nodes.Clear
    SMnu = GetSystemMenu(SelectedNodeHwnd, 0)
    BMnu = GetMenu(SelectedNodeHwnd)
    If SMnu Then
        MnuTree.Nodes.Add , , "SMenu", "System Menu"
        mchild MnuTree, SMnu, "SMenu", "S:"
    End If
    If BMnu Then
        MnuTree.Nodes.Add , , "BMenu", "Menu Bar"
        mchild MnuTree, BMnu, "BMenu", "B:"
    End If
    MnuTree.Nodes.item(1).Selected = True
    LockWindowUpdate 0
    FillListWithMenuItems MnuList

End Sub

Private Sub TreeMenuWindowStyles_Click()

    TaskMenuID = 1
    TaskTree.Enabled = False
    LockWindowUpdate FrameMWS.hwnd
    'i made a function to easily setup my User interface on controls
    SetProps FrameMWS, 0, 0, 9710, TaskTree.Height, True
    SetProps FrameWSX, 0, 6120, 3590, FrameMWS.Height, True
    SetProps FrameWS, 0, 2530, 3590, FrameMWS.Height, True
    SetProps WSXList, 240, 120, 3375, FrameWSX.Height - 200, True
    SetProps WSList, 240, 120, 3375, FrameWS.Height - 200, True
    SetProps WSButton(0), 360, bLeft, bWidth, bHeight, True, , "Refresh"
    Load WSButton(1)
    SetProps WSButton(1), FrameMWS.Height - 600, bLeft, bWidth, bHeight, True, , "Close This Menu"

    WSList.Clear
    WSXList.Clear
    AddToList WSList
    AddToListX WSXList
    ListGetStyles WSList, SelectedNodeHwnd
    ListGetStylesX WSXList, SelectedNodeHwnd
    LockWindowUpdate 0

End Sub

Private Sub WSButton_Click(Index As Integer)
'TaskMenuId Windows
'1 Window Styles
    Select Case Index
      Case 0 'refresh
        Unload WSButton(1)
        TreeMenuWindowStyles_Click
      Case 1 'close
        TaskMenuID = 0
        FrameMWS.Visible = False
        TaskTree.Enabled = True
        Unload WSButton(1)
    End Select

End Sub

Private Sub WSList_ItemCheck(item As Integer)

    If AuthorMode Then
        SetWS SelectedNodeHwnd, item, WSList.Selected(item)
    End If

End Sub

Private Sub WSXList_ItemCheck(item As Integer)

    If AuthorMode Then
        SetWSX SelectedNodeHwnd, item, WSXList.Selected(item)
    End If

End Sub
