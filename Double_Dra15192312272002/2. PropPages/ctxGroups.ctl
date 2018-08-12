VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl ctxGroups 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6576
   ScaleHeight     =   4560
   ScaleWidth      =   6576
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   2520
      Top             =   0
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Icons And Graphics (*.ico;*.bmp;*.gif;*.jpg)|*.ico;*.bmp;*.gif;*.jpg|All files (*.*)|*.*"
      Flags           =   4
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontStrikeThru  =   -1  'True
      FontUnderLine   =   -1  'True
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   285
      Left            =   1146
      Picture         =   "ctxGroups.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   84
      Width           =   300
   End
   Begin VB.CommandButton cmdRemove 
      Height          =   285
      Left            =   1476
      Picture         =   "ctxGroups.ctx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   84
      Width           =   300
   End
   Begin VB.CommandButton cmdRename 
      Height          =   285
      Left            =   1806
      Picture         =   "ctxGroups.ctx":0294
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   84
      Width           =   300
   End
   Begin VB.PictureBox picControls 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4464
      Left            =   2268
      ScaleHeight     =   4464
      ScaleWidth      =   4212
      TabIndex        =   17
      Top             =   0
      Width           =   4212
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         Height          =   285
         Left            =   2610
         TabIndex        =   30
         Top             =   4050
         Width           =   1365
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   285
         Left            =   1170
         TabIndex        =   29
         Top             =   4050
         Width           =   1365
      End
      Begin VB.CommandButton cmdLargeCopy 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1848
         Picture         =   "ctxGroups.ctx":03DE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2856
         Width           =   300
      End
      Begin VB.CommandButton cmdSmallCopy 
         Height          =   285
         Left            =   1848
         Picture         =   "ctxGroups.ctx":0528
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2184
         Width           =   300
      End
      Begin VB.CommandButton cmdLargeIcon 
         Height          =   285
         Left            =   2508
         Picture         =   "ctxGroups.ctx":0672
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2856
         Width           =   300
      End
      Begin VB.CommandButton cmdSmallIcon 
         Height          =   285
         Left            =   2508
         Picture         =   "ctxGroups.ctx":07BC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2184
         Width           =   300
      End
      Begin VB.TextBox txtIndex 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   1176
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   1764
         Width           =   1440
      End
      Begin VB.TextBox txtKey 
         Height          =   288
         Left            =   1176
         TabIndex        =   7
         Top             =   924
         Width           =   2952
      End
      Begin VB.TextBox txtTag 
         Height          =   288
         Left            =   1176
         TabIndex        =   8
         Top             =   1344
         Width           =   2952
      End
      Begin VB.TextBox txtTooltip 
         Height          =   288
         Left            =   1176
         TabIndex        =   6
         ToolTipText     =   "Multiline: use \n for newline separator"
         Top             =   504
         Width           =   2952
      End
      Begin VB.TextBox txtCaption 
         Height          =   288
         Left            =   1176
         TabIndex        =   5
         ToolTipText     =   "Multiline: use \n for newline separator"
         Top             =   84
         Width           =   2952
      End
      Begin VB.CommandButton cmdSmallClear 
         Height          =   285
         Left            =   2168
         Picture         =   "ctxGroups.ctx":0906
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2184
         Width           =   300
      End
      Begin VB.CommandButton cmdLargeClear 
         Height          =   285
         Left            =   2168
         Picture         =   "ctxGroups.ctx":0E90
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2856
         Width           =   300
      End
      Begin VB.PictureBox picSmall 
         AutoRedraw      =   -1  'True
         Height          =   600
         Left            =   1176
         ScaleHeight     =   552
         ScaleWidth      =   552
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2184
         Width           =   600
         Begin VB.Image imgSmall 
            Height          =   348
            Left            =   84
            MousePointer    =   15  'Size All
            Top             =   84
            Width           =   348
         End
      End
      Begin VB.PictureBox picLarge 
         AutoRedraw      =   -1  'True
         Height          =   600
         Left            =   1176
         ScaleHeight     =   552
         ScaleWidth      =   552
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2856
         Width           =   600
         Begin VB.Image imgLarge 
            Height          =   348
            Left            =   84
            MousePointer    =   15  'Size All
            Top             =   84
            Width           =   348
         End
      End
      Begin VB.OptionButton optIconsSmall 
         Caption         =   "Small Icons"
         Height          =   264
         Left            =   1176
         TabIndex        =   15
         Top             =   3612
         Value           =   -1  'True
         Width           =   1356
      End
      Begin VB.OptionButton optIconsLarge 
         Caption         =   "Large Icons"
         Height          =   264
         Left            =   2604
         TabIndex        =   16
         Top             =   3612
         Width           =   1524
      End
      Begin VB.Label Label7 
         Caption         =   "Key:"
         Height          =   264
         Left            =   84
         TabIndex        =   28
         Top             =   924
         Width           =   1188
      End
      Begin VB.Label Label6 
         Caption         =   "Tag:"
         Height          =   264
         Left            =   84
         TabIndex        =   27
         Top             =   1344
         Width           =   1188
      End
      Begin VB.Label Label5 
         Caption         =   "TooltipText:"
         Height          =   264
         Left            =   84
         TabIndex        =   26
         Top             =   504
         Width           =   1188
      End
      Begin VB.Label Label4 
         Caption         =   "Index:"
         Height          =   264
         Left            =   84
         TabIndex        =   25
         Top             =   1764
         Width           =   1188
      End
      Begin VB.Label Label3 
         Caption         =   "Large Icon:"
         Height          =   264
         Left            =   84
         TabIndex        =   24
         Top             =   2856
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Small Icon:"
         Height          =   264
         Left            =   84
         TabIndex        =   23
         Top             =   2184
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Caption:"
         Height          =   264
         Left            =   84
         TabIndex        =   22
         Top             =   84
         Width           =   1020
      End
      Begin VB.Label Label8 
         Caption         =   "Icons Style:"
         Height          =   264
         Left            =   84
         TabIndex        =   21
         Top             =   3612
         Width           =   1020
      End
   End
   Begin MSComctlLib.TreeView trvGroups 
      Height          =   4044
      Left            =   84
      TabIndex        =   1
      Top             =   420
      Width           =   2028
      _ExtentX        =   3577
      _ExtentY        =   7154
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label10 
      Caption         =   "&Buttons:"
      Height          =   312
      Left            =   84
      TabIndex        =   0
      Top             =   84
      Width           =   1152
   End
End
Attribute VB_Name = "ctxGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqweto@myrealbox.com)
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "ctxGroups"

'=========================================================================
' Events
'=========================================================================

Event Changed()

'=========================================================================
' Constants and variables
'=========================================================================

Private Const CAP_MSG               As String = "Groups Property Page"
Private Const STR_NEW_GROUP         As String = "New Group"
Private Const STR_NEW_ITEM          As String = "New Item"
Private Const STR_ROOT              As String = "Root"
Private Const STR_BUTTONBAR         As String = "Button Bar"
Private Const MSG_COPY_SMALL_TO_LARGE   As String = "Do you want to copy small icon to large icon?"
Private Const MSG_COPY_LARGE_TO_SMALL   As String = "Do you want to copy large icon to small icon?"

Private m_oControl              As Object
Private m_oGroups               As cButton
Private m_oSel                  As cButton
Private m_bInSet                As Boolean
Private m_bDrag                 As Boolean
Private m_sX                    As Single
Private m_sY                    As Single

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunc As String)
    PushError sFunc, MODULE_NAME
    PopRaiseError
End Sub

Private Function ShowError(sFunc As String) As VbMsgBoxResult
    PushError sFunc, MODULE_NAME
    ShowError = PopShowError(CAP_MSG)
End Function

'=========================================================================
' Properties
'=========================================================================

Property Let Changed(ByVal bValue As Boolean)
    RaiseEvent Changed
End Property

'=========================================================================
' Methods
'=========================================================================

Private Sub pvLoadGroups()
    Const FUNC_NAME     As String = "pvLoadGroups"
    Dim oGrp            As cButton
    Dim oItm            As cButton
    Dim sKey            As String
    Dim bFocused        As Boolean
        
    On Error GoTo EH
    '--- save selected node
    sKey = trvGroups.SelectedItem.Key
    '--- hide for faster access
    bFocused = ActiveControl Is trvGroups
    trvGroups.Visible = False
    '--- do load
    With trvGroups.Nodes
        .Clear
        .Add(, , STR_ROOT, STR_BUTTONBAR).Expanded = True
        For Each oGrp In m_oGroups.GroupItems
            .Add(STR_ROOT, tvwChild, "g" & oGrp.Index, oGrp.Caption).Expanded = True
            For Each oItm In oGrp.GroupItems
                .Add "g" & oGrp.Index, tvwChild, "i" & oItm.Index & "g" & oGrp.Index, Replace(oItm.Caption, vbCrLf, "\n")
            Next
        Next
    End With
    '--- restore selected node
    On Error Resume Next
    Set trvGroups.SelectedItem = trvGroups.Nodes(sKey)
    Set m_oSel = pvGetButton(trvGroups.SelectedItem)
    On Error GoTo EH
    pvFillControls m_oSel
    '--- restore visibility
    trvGroups.Visible = True
    If bFocused Then
        trvGroups.SetFocus
    End If
    Exit Sub
EH:
    PushError FUNC_NAME, MODULE_NAME
    trvGroups.Visible = True
    PopRaiseError
End Sub

Private Function pvGetButton(oNode As Node) As cButton
    Const FUNC_NAME     As String = "pvGetButton"
    Dim sKey            As String
    
    On Error GoTo EH
    If Not oNode Is Nothing Then
        sKey = oNode.Key
        If sKey = STR_ROOT Then
            Set pvGetButton = m_oGroups
        ElseIf Left(sKey, 1) = "g" Then
            Set pvGetButton = m_oGroups.GroupItems(Val(Mid(sKey, 2)))
        Else
            Set pvGetButton = m_oGroups.GroupItems(C2Lng(Mid(trvGroups.SelectedItem.Parent.Key, 2)))
            Set pvGetButton = pvGetButton.GroupItems(Val(Mid(sKey, 2)))
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvFillControls(oSel As cButton)
    Const FUNC_NAME     As String = "pvFillControls"
    
    On Error GoTo EH
    If oSel Is Nothing Then
        picControls.Visible = False
    Else
        m_bInSet = True
        picControls.Visible = False
        With oSel
            picControls.Visible = .Class <> ucsBtnClassControl
            txtCaption = Replace(.Caption, vbCrLf, "\n")
            txtTooltip = Replace(.ToolTipText, vbCrLf, "\n")
            txtKey = .Key
            txtTag = C2Str(.Tag)
            txtIndex = .Index
            Set imgSmall = .SmallIcon
            Set imgLarge = .LargeIcon
            pvCenterIcons
            optIconsLarge.Enabled = .Class = ucsBtnClassGroup
            optIconsSmall.Enabled = optIconsLarge.Enabled
            optIconsLarge = .IconsType = ucsIcsLargeIcons
            optIconsSmall = Not optIconsLarge
            chkEnabled.Enabled = .Class <> ucsBtnClassControl
            chkVisible.Enabled = chkEnabled.Enabled
            chkEnabled = IIf(.Enabled, vbChecked, vbUnchecked)
            chkVisible = IIf(.Visible, vbChecked, vbUnchecked)
        End With
        picControls.Visible = True
        m_bInSet = False
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvUploadButton()
    Const FUNC_NAME     As String = "pvUploadButton"
    
    On Error GoTo EH
    If Not m_oSel Is Nothing Then
        With m_oSel
            Select Case .Class
            Case ucsBtnClassGroup
                trvGroups.Nodes("g" & .Index).Text = txtCaption
            Case ucsBtnClassItem
                trvGroups.Nodes("i" & .Index & "g" & .Parent.Index).Text = txtCaption
            End Select
            .Caption = Replace(txtCaption, "\n", vbCrLf)
            .ToolTipText = Replace(txtTooltip, "\n", vbCrLf)
            .Key = txtKey
            .Tag = C2Str(txtTag)
            Set .SmallIcon = imgSmall.Picture
            Set .LargeIcon = imgLarge.Picture
            If .Class = ucsBtnClassGroup Then
                .IconsType = IIf(optIconsSmall, ucsIcsSmallIcons, ucsIcsLargeIcons)
            End If
            .Enabled = (chkEnabled = vbChecked)
            .Visible = (chkVisible = vbChecked)
        End With
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvCenterIcons()
    On Error Resume Next
    imgSmall.Move (picSmall.ScaleWidth - imgSmall.Width) \ 2, (picSmall.ScaleHeight - imgSmall.Height) \ 2
    imgLarge.Move (picLarge.ScaleWidth - imgLarge.Width) \ 2, (picLarge.ScaleHeight - imgLarge.Height) \ 2
    picLarge.Visible = False
    picLarge.Visible = True
    picSmall.Visible = False
    picSmall.Visible = True
End Sub

Private Sub pvSetControlSelected(oSel As cButton)
    Const FUNC_NAME     As String = "pvSetControlSelected"
    
    '--- silence on errors: groups/items might've already not been applyed
'    On Error GoTo EH
    On Error Resume Next
    If Not oSel Is Nothing Then
        With oSel
            If .Class = ucsBtnClassGroup Then
                m_oControl.Groups(.Index).Selected = True
            ElseIf .Class = ucsBtnClassItem Then
                m_oControl.Groups(.Parent.Index)(.Index).Selected = True
            End If
        End With
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'=========================================================================
' Control events
'=========================================================================

Public Sub SelectionChanged(SelectedControls As Object)
    Const FUNC_NAME     As String = "SelectionChanged"
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    If Not m_oControl Is SelectedControls(0) Then
        Set m_oControl = SelectedControls(0)
        Set m_oGroups = New cButton
        m_oGroups.Class = ucsBtnClassControl
        m_oGroups.GroupItems.Contents = m_oControl.Groups.Contents
        '--- make sure somthing is selected
        trvGroups.Nodes.Clear
        trvGroups.Nodes.Add , , STR_ROOT, ""
        Set trvGroups.SelectedItem = trvGroups.Nodes(1)
        pvLoadGroups
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub ApplyChanges()
    Const FUNC_NAME     As String = "ApplyChanges"
    
    On Error GoTo EH
    pvUploadButton
    Set m_oControl.Groups = m_oGroups.GroupItems
    pvSetControlSelected m_oSel
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub cmdLargeCopy_Click()
    If MsgBox(MSG_COPY_LARGE_TO_SMALL, vbQuestion + vbYesNo, CAP_MSG) = vbYes Then
        Set imgSmall.Picture = imgLarge.Picture
        Changed = True
        pvCenterIcons
    End If
End Sub

Private Sub cmdSmallCopy_Click()
    If MsgBox(MSG_COPY_SMALL_TO_LARGE, vbQuestion + vbYesNo, CAP_MSG) = vbYes Then
        Set imgLarge.Picture = imgSmall.Picture
        Changed = True
        pvCenterIcons
    End If
End Sub

Private Sub trvGroups_AfterLabelEdit(Cancel As Integer, NewString As String)
    Const FUNC_NAME     As String = "trvGroups_AfterLabelEdit"
    
    On Error GoTo EH
    Set m_oSel = pvGetButton(trvGroups.SelectedItem)
    m_oSel.Caption = NewString
    txtCaption = NewString
    Changed = True
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub trvGroups_BeforeLabelEdit(Cancel As Integer)
    '--- prevent caption edit of root
    On Error Resume Next
    If trvGroups.SelectedItem.Index = 1 Then
        Cancel = 1
    End If
End Sub

Private Sub trvGroups_DblClick()
    Const FUNC_NAME     As String = "trvGroups_DblClick"
    
    On Error GoTo EH
    trvGroups.SelectedItem.Expanded = True
    trvGroups.StartLabelEdit
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub trvGroups_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "trvGroups_KeyDown"
    
    On Error GoTo EH
    If KeyCode = vbKeyF2 And Shift = 0 Then
        trvGroups.SetFocus
        trvGroups.StartLabelEdit
    End If
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub trvGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    Const FUNC_NAME     As String = "trvGroups_NodeClick"
    
    On Error GoTo EH
    pvUploadButton
    Set m_oSel = pvGetButton(trvGroups.SelectedItem)
    pvFillControls m_oSel
    pvSetControlSelected m_oSel
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub cmdAdd_Click()
    Const FUNC_NAME     As String = "cmdAdd_Click"
    Dim oBtn            As cButton
    
    On Error GoTo EH
    If m_oSel Is Nothing Then
        Exit Sub
    End If
    pvUploadButton
    If m_oSel.Class = ucsBtnClassItem Then
        Set oBtn = m_oSel.Parent.GroupItems.Add(STR_NEW_ITEM & m_oSel.Parent.GroupItems.Count, , , , , m_oSel.Index)
    Else
        Set oBtn = m_oSel.GroupItems.Add(IIf(m_oSel.Class = ucsBtnClassControl, STR_NEW_GROUP, STR_NEW_ITEM) & m_oSel.GroupItems.Count)
    End If
    oBtn.ToolTipText = ""
    pvLoadGroups
    Changed = True
    With trvGroups
        If oBtn.Class = ucsBtnClassItem Then
            Set .SelectedItem = .Nodes("i" & oBtn.Index & "g" & oBtn.Parent.Index)
        Else
            Set .SelectedItem = .Nodes("g" & oBtn.Index)
        End If
        .SetFocus
        .StartLabelEdit
    End With
    Set m_oSel = pvGetButton(trvGroups.SelectedItem)
    pvFillControls m_oSel
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub cmdRemove_Click()
    Const FUNC_NAME     As String = "cmdRemove_Click"
    
    On Error GoTo EH
    If m_oSel Is Nothing Then
        Exit Sub
    End If
    If m_oSel.Class = ucsBtnClassItem Or _
            m_oSel.Class = ucsBtnClassGroup Then
        m_oSel.Parent.GroupItems.Remove m_oSel.Index
        Set m_oSel = Nothing
        pvLoadGroups
        Changed = True
    End If
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub cmdRename_Click()
    Const FUNC_NAME     As String = "cmdRename_Click"
    
    On Error GoTo EH
    trvGroups.SetFocus
    trvGroups.StartLabelEdit
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub txtCaption_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtTooltip_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtKey_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtTag_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub optIconsLarge_Click()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub optIconsSmall_Click()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub chkEnabled_Click()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub chkVisible_Click()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub cmdSmallIcon_Click()
    On Error GoTo EH_Cancel
    comDlg.ShowOpen
    Set imgSmall = LoadPicture(comDlg.FileName)
    Changed = True
    DoEvents
    pvCenterIcons
EH_Cancel:
End Sub

Private Sub cmdSmallClear_Click()
    On Error Resume Next
    Set imgSmall = Nothing
    Changed = True
End Sub

Private Sub cmdLargeIcon_Click()
    On Error GoTo EH_Cancel
    comDlg.ShowOpen
    Set imgLarge = LoadPicture(comDlg.FileName)
    Changed = True
    DoEvents
    pvCenterIcons
EH_Cancel:
End Sub

Private Sub cmdLargeClear_Click()
    On Error Resume Next
    Set imgLarge = Nothing
    Changed = True
End Sub

Private Sub imgSmall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    m_bDrag = True
    m_sX = X: m_sY = Y
End Sub

Private Sub imgSmall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If m_bDrag Then
        With imgSmall
            .Move .Left + (X - m_sX), .Top + (Y - m_sY)
        End With
    End If
End Sub

Private Sub imgSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    imgSmall_MouseMove Button, Shift, X, Y
    With imgSmall
        If .Width < picSmall.ScaleWidth Then
            .Left = (picSmall.ScaleWidth - .Width) \ 2
        Else
            If .Left < picSmall.ScaleWidth - .Width Then
                .Left = picSmall.ScaleWidth - .Width
            End If
            If .Left > 0 Then
                .Left = 0
            End If
        End If
        If .Height < picSmall.ScaleHeight Then
            .Top = (picSmall.ScaleHeight - .Height) \ 2
        Else
            If .Top < picSmall.ScaleHeight - .Height Then
                .Top = picSmall.ScaleHeight - .Height
            End If
            If .Top > 0 Then
                .Top = 0
            End If
        End If
    End With
    m_bDrag = False
End Sub


Private Sub imgLarge_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    m_bDrag = True
    m_sX = X: m_sY = Y
End Sub

Private Sub imgLarge_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If m_bDrag Then
        With imgLarge
            .Move .Left + (X - m_sX), .Top + (Y - m_sY)
        End With
    End If
End Sub

Private Sub imgLarge_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    imgLarge_MouseMove Button, Shift, X, Y
    With imgLarge
        If .Width < picLarge.ScaleWidth Then
            .Left = (picLarge.ScaleWidth - .Width) \ 2
        Else
            If .Left < picLarge.ScaleWidth - .Width Then
                .Left = picLarge.ScaleWidth - .Width
            End If
            If .Left > 0 Then
                .Left = 0
            End If
        End If
        If .Height < picLarge.ScaleHeight Then
            .Top = (picLarge.ScaleHeight - .Height) \ 2
        Else
            If .Top < picLarge.ScaleHeight - .Height Then
                .Top = picLarge.ScaleHeight - .Height
            End If
            If .Top > 0 Then
                .Top = 0
            End If
        End If
    End With
    m_bDrag = False
End Sub



