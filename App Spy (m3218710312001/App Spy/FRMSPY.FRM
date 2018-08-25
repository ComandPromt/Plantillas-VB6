VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSpy 
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   9540
   Icon            =   "FRMSPY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      TabCaption(0)   =   "Active Tasks"
      TabPicture(0)   =   "FRMSPY.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ImageList1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwTask"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pIcon"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRact"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Active Windows"
      TabPicture(1)   =   "FRMSPY.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwParent"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lvwChild"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkCap"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdEnum"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "HScroll1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Startup Tasks"
      TabPicture(2)   =   "FRMSPY.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "imgStart"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lvwRun"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdRReg"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdDelKey"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdRegedit"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "pic2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdWin"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdSys"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.CommandButton cmdSys 
         Caption         =   "System.ini"
         Height          =   255
         Left            =   -70920
         TabIndex        =   15
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdWin 
         Caption         =   "Win.ini"
         Height          =   255
         Left            =   -72000
         TabIndex        =   14
         Top             =   5160
         Width           =   975
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         LargeChange     =   30
         Left            =   -74880
         Max             =   9150
         Min             =   1
         SmallChange     =   15
         TabIndex        =   13
         Top             =   4920
         Value           =   4575
         Width           =   9255
      End
      Begin VB.PictureBox pic2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -66960
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   12
         Top             =   5160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton cmdEnum 
         Caption         =   "E&numerate"
         Height          =   375
         Left            =   -66720
         TabIndex        =   11
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CheckBox chkCap 
         Caption         =   "&Show only windows with captions."
         Height          =   195
         Left            =   -74880
         TabIndex        =   8
         Top             =   5280
         Width           =   2775
      End
      Begin VB.CommandButton cmdRegedit 
         Caption         =   "Regedit"
         Height          =   375
         Left            =   -66600
         TabIndex        =   7
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdDelKey 
         Caption         =   "Remove Entry"
         Height          =   375
         Left            =   -73800
         TabIndex        =   6
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdRReg 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdRact 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   8160
         TabIndex        =   2
         ToolTipText     =   "Refresh avtive task list."
         Top             =   5160
         Width           =   1215
      End
      Begin VB.PictureBox pIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7800
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   1
         Top             =   5160
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSComctlLib.ListView lvwTask 
         Height          =   4575
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Display name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Path"
            Object.Width           =   8148
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Priority"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Usage"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Threads"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6840
         Top             =   5040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lvwRun 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "key"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Command"
            Object.Width           =   12347
         EndProperty
      End
      Begin MSComctlLib.ListView lvwChild 
         Height          =   4455
         Left            =   -70200
         TabIndex        =   9
         ToolTipText     =   "Child windows"
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7858
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwParent 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   10
         ToolTipText     =   "Parent windows"
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7858
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ImageList imgStart 
         Left            =   -67680
         Top             =   5040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Menu mnuActTsk 
      Caption         =   "ActTsk"
      Visible         =   0   'False
      Begin VB.Menu mnuProp 
         Caption         =   "File Properties"
      End
      Begin VB.Menu mnuMod 
         Caption         =   "Used Modules"
      End
      Begin VB.Menu mnuStr 
         Caption         =   "View"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEndTsk 
         Caption         =   "End Task"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu Show 
         Caption         =   "&Show Window using ShowWindow API"
      End
      Begin VB.Menu Show_BWTT 
         Caption         =   "Show &Winsow using BringWindowToTop API"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu Max 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu Min 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu Restore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu Hide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close this Window"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "menu2"
      Visible         =   0   'False
      Begin VB.Menu BnClick 
         Caption         =   "&Click"
      End
   End
End
Attribute VB_Name = "frmSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Clicked As ListItem
Dim NumOfProcess As Long
Dim objActiveProcess As GetPro
Private R As New cReg
Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "There is already another instance of this program running.", vbCritical, "Error"
    Unload Me
Else
    Set objActiveProcess = New GetPro
    frmSpy.Caption = App.ProductName & "  " & App.Major & "." & App.Minor
    LoadActTsk
    LoadActWin
    LoadStartUp
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case vbFormControlMenu: cMem 'Pushing "X" requested close.
        Case vbFormCode: cMem 'Code requested close.
        Case vbAppWindows: cMem 'Windows Shutting down: requesting all programs close.
        Case vbAppTaskManager: cMem 'Task Manager requested close
    End Select
End Sub
'
'*******Active Tasks*******
Private Sub cmdRact_Click()
LoadActTsk
End Sub
Private Sub LoadActTsk()
On Error Resume Next
Dim Ret
Set lvwTask.SmallIcons = Nothing
lvwTask.ListItems.Clear
ImageList1.ListImages.Clear
NumOfProcess = objActiveProcess.GetActiveProcess
ImageList1.ListImages.Add , , pIcon.Image
Set lvwTask.SmallIcons = ImageList1
For i = 1 To NumOfProcess
    Ret = GetIcon(objActiveProcess.exePath(i), pIcon)
    ImageList1.ListImages.Add i + 1, , pIcon.Image
    lvwTask.ListItems.Add , , Ret, , i + 1
    lvwTask.Refresh
    With lvwTask.ListItems.Item(i)
      .ListSubItems.Add , , objActiveProcess.exePath(i)
      If fEnumWindows(objActiveProcess.ProcessID(i)) = 0 Then
          .ListSubItems.Add , , "Running"
      Else
          .ListSubItems.Add , , "Frozen"
          .ForeColor = vbRed
          .ListSubItems.Item(1).ForeColor = vbRed
          .ListSubItems.Item(2).ForeColor = vbRed
      End If
      'Determine process priority based on it's flag number.
      Select Case objActiveProcess.PriClassBase(i)
          Case "24": .ListSubItems.Add , , "RealTime"
          Case "13": .ListSubItems.Add , , "High"
          Case "8": .ListSubItems.Add , , "Normal"
          Case "4": .ListSubItems.Add , , "Idle"
          Case Else
      End Select
      .ListSubItems.Add , , objActiveProcess.Usage(i)
      .ListSubItems.Add , , objActiveProcess.Threads(i)
    End With
Next i
End Sub
Private Sub HScroll1_Change()
    lvwParent.Width = HScroll1.Value
    lvwChild.Width = (frmSpy.Width - 510) - HScroll1.Value
    lvwChild.Left = lvwParent.Width + 225
End Sub

Private Sub mnuProp_Click()
ShowProperties objActiveProcess.exePath(Clicked.Index), Me.hwnd
End Sub
Private Sub mnuMod_Click() 'Shows what dll(s) and or ocx(s) are used by the selected program.
Dim procStr() As String, modStr As String, n As Integer
On Error GoTo Err
procStr() = GetProcessModules(objActiveProcess.ProcessID(Clicked.Index))
    For i = LBound(procStr()) To UBound(procStr())
        modStr = modStr & vbCrLf & procStr(i)
        i = i + 1
    Next i
MsgBox "Used process modules for: " & Clicked.Text & modStr, vbInformation, Clicked.Text
Err:
    Exit Sub
End Sub
Private Sub mnuStr_Click()
frmStr.rtbStr.Text = ""
frmStr.rtbStr.LoadFile objActiveProcess.exePath(Clicked.Index)
frmStr.Caption = lvwTask.ListItems.Item(Clicked.Index).Text
frmStr.Icon = Me.Icon
frmStr.Show
End Sub
Private Sub mnuEndTsk_Click()
Dim lProcess As Long
Dim lReturn As Long
Dim Ret As VbMsgBoxResult
If Clicked.Text = "AppSpy.exe" Then
    MsgBox "You realize that closing me this way is redundant." & vbCrLf & "Have you tried pushing the little X?", vbCritical, "Moronic Error"
    Exit Sub
End If
If MsgBox("Are you sure you want to terminate this?" & vbCrLf & vbCrLf & UCase$(Clicked.Text), vbExclamation + vbYesNo, "Warning") = vbYes Then
    lProcess = OpenProcess(&H1F0FFF, 0&, objActiveProcess.ProcessID(Clicked.Index))
    lReturn = TerminateProcess(lProcess, 0&)
    DoEvents 'Give time for process to unload, then refresh active processes
    LoadActTsk
End If
End Sub
Private Sub lvwTask_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set Clicked = Item
End Sub
Private Sub lvwTask_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton And lvwTask.ListItems.count > 0 Then PopupMenu mnuActTsk
End Sub
'*******END Active Tasks*******
'
'
'*******Active Windows*******
Private Sub LoadActWin()
lvwParent.View = lvwReport
With lvwParent.ColumnHeaders
    .Add , , "Handle", 750
    .Add , , "Class Name", 1500
    .Add , , "Text", 3000
End With
VCount = 1
lvwChild.View = lvwReport
With lvwChild.ColumnHeaders
    .Add , , "Handle", 750
    .Add , , "Class Name", 1500
    .Add , , "Text", 2500
    .Add , , "Password Field", 1400
End With
ICount = 1
Options.Visible = False
End Sub
Private Sub BnClick_Click() 'Click selected item if it is a button
    SendMessage Val(lvwChild.SelectedItem), BM_CLICK, 0, 0
End Sub
Private Sub cmdEnum_Click()
    cmdEnum.Caption = "&Refresh"
    lvwParent.ListItems.Clear
    lvwChild.ListItems.Clear
    lvwParent.GridLines = True
    VCount = 1
    EnumWindows AddressOf WndEnumProc, lvwParent
End Sub
Private Sub Close_Click() 'close window
On Error Resume Next
Dim lhwnd As Long
    lhwnd = Val(lvwParent.SelectedItem)
    SendMessage lhwnd, WM_CLOSE, 0, 0
End Sub
Private Sub Hide_Click()
    ShowWindow Val(lvwParent.SelectedItem), SW_HIDE
End Sub
Private Sub Max_Click()
    ShowWindow Val(lvwParent.SelectedItem), SW_MAXIMIZE
End Sub
Private Sub Min_Click()
    ShowWindow Val(lvwParent.SelectedItem), SW_MINIMIZE
End Sub
Private Sub Restore_Click()
    ShowWindow Val(lvwParent.SelectedItem), SW_RESTORE
End Sub
Private Sub Show_BWTT_Click()
Dim lhwnd As Long
    lhwnd = Val(lvwParent.SelectedItem)
    BringWindowToTop lhwnd
End Sub
Private Sub Show_Click() 'show window
On Error Resume Next
Dim lhwnd As Long
    lhwnd = Val(lvwParent.SelectedItem)
    ShowWindow lhwnd, SW_SHOW
End Sub
Private Sub lvwParent_Click()
    GotoChild
End Sub
Private Sub lvwParent_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then GotoChild
                'So that you are able to see child windows easily by
                'scrolling through up-down arrow keys instead of
                'clicking the parent window handle every time.
End Sub
Private Sub GotoChild()
On Error Resume Next
Dim Num As Long
    Num = Val(lvwParent.SelectedItem)
    lvwChild.ListItems.Clear
    lvwChild.GridLines = True
    ICount = 1
    EnumChildWindows Num, AddressOf WndEnumChildProc, lvwChild
End Sub
Private Sub lvwParent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And lvwParent.ListItems.count > 0 Then PopupMenu Options
End Sub
Private Sub lvwChild_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And lvwChild.ListItems.count > 0 Then PopupMenu menu2
End Sub
'*******END Active Windows*******
'
'
'*******Startup Tasks*******
Private Sub cmdRReg_Click() 'Refresh
LoadStartUp
End Sub
Private Sub LoadStartUp()
Dim lrtn As Long, X As Integer, Y As Integer
Dim Str1 As String, Item As ListItem
lKeys
lvwRun.ListItems.Clear
cmdDelKey.Enabled = False
    For X = 1 To 4
        lrtn = R.GetKey(hKeys(X).Class, hKeys(X).Key, MainKeys())
        If (lrtn > 0) Then
            For Y = 1 To Int(lrtn)
                Set Item = lvwRun.ListItems.Add()
                If Y = 1 Then
                    Item.Text = hKeys(X).Name
                    Item.Tag = hKeys(X).Name
                Else
                    Item.Text = ""
                    Item.Tag = hKeys(X).Name
                End If
                If MainKeys(Y) = "" Then
                    Item.SubItems(1) = "Empty..."
                Else
                    Str1 = R.qVal(hKeys(X).Class, hKeys(X).Key, MainKeys(Y))
                    Item.SubItems(1) = MainKeys(Y)
                    Item.SubItems(2) = Str1
                End If
            Next
            Set Item = lvwRun.ListItems.Add()
        Else
            Set Item = lvwRun.ListItems.Add()
            Item.Text = hKeys(X).Name
            Item.SubItems(1) = "Empty..."
            Set Item = lvwRun.ListItems.Add()
        End If
    Next
With lvwRun.ListItems
    n = .count
    .Add (n), , ""
    n = n + 1
    .Add (n), , "Win.ini"
    .Item(n).ListSubItems.Add , , "Load"
    .Item(n).ListSubItems.Add , , ReadINI("Windows", "load", "Win.ini")
    .Item(n).Tag = "Load"
    n = n + 1
    .Add (n), , ""
    .Item(n).ListSubItems.Add , , "Run"
    .Item(n).ListSubItems.Add , , ReadINI("Windows", "run", "Win.ini")
    .Item(n).Tag = "Run"
    n = n + 1
    .Add (n), , ""
    .Item(n).ListSubItems.Add , , ""
    .Item(n).ListSubItems.Add , , ""
    n = n + 1
    .Add (n), , "System.ini"
    .Item(n).ListSubItems.Add , , "Shell"
    .Item(n).ListSubItems.Add , , ReadINI("boot", "shell", "System.ini")
    .Item(n).Tag = "Shell"
End With
End Sub
Private Sub cmdDelKey_Click()
Dim iNum As Long, iTag As String, hKey As Long, nKey As String, cver As String
cver = "Software\Microsoft\Windows\CurrentVersion\"
iNum = lvwRun.SelectedItem.Index
iTag = lvwRun.ListItems.Item(iNum).Tag
nKey = lvwRun.ListItems.Item(iNum).ListSubItems(1).Text
If MsgBox("Are you sure you want to remove this?" & vbCrLf & vbCrLf & UCase$(nKey), vbExclamation + vbYesNo, "Warning") = vbYes Then
    Select Case iTag
        Case "User_Run": R.DelVal &H80000001, cver & "Run", nKey
        Case "User_RunServices": R.DelVal &H80000001, cver & "Runservices", nKey
        Case "Machine_Run": R.DelVal &H80000002, cver & "Run", nKey
        Case "Machine_RunServices": R.DelVal &H80000002, cver & "Runservices", nKey
        Case "Load": WriteINI "Windows", "load", "Win.ini", ""
        Case "Run": WriteINI "Windows", "Run", "Win.ini", ""
        Case "Shell": If MsgBox("Revert to the default Shell, explorer.exe?", vbExclamation + vbYesNo, _
                "Warning") = vbYes Then WriteINI "boot", "shell", "system.ini", "explorer.exe"
        Case Else
    End Select
    LoadStartUp
Else
    Exit Sub
End If
End Sub
Private Sub cmdWin_Click()
Shell WinDir & "\notepad.exe " & WinDir & "\Win.ini"
End Sub
Private Sub cmdSys_Click()
Shell WinDir & "\notepad.exe " & WinDir & "\System.ini"
End Sub
Private Sub cmdRegedit_Click()
Shell "c:\Windows\regedit.exe"
End Sub
Private Sub lvwRun_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim iNum As Long, iTag As String, iTxt As String
iNum = lvwRun.SelectedItem.Index
iTxt = lvwRun.ListItems.Item(iNum).ListSubItems(2).Text
If iTxt = "" Or iTxt = "Empty..." Then cmdDelKey.Enabled = False Else cmdDelKey.Enabled = True
End Sub
'*******END Startup Tasks*******
'
Function cMem() 'Clean up memory.
Set objActiveProcess = Nothing
Set R = Nothing
End Function
