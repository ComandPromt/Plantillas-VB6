VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Dynamic TreeView, and ListView Sample. ( Coded by Chris Hatton. www.chris.hatton.com)"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10275
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3600
      Top             =   2760
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   953
      ButtonWidth     =   1693
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Record"
            Object.ToolTipText     =   "Add Record (F1)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit Record"
            Object.ToolTipText     =   "Edit Record (F2)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Del Record"
            Object.ToolTipText     =   "Delete Record (F3)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View Table"
            Object.ToolTipText     =   "Export Selection to ListView Control (F4)"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Comdiag 
      Left            =   3600
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1104
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1556
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":16B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":25CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2726
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":31F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":48DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":53A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView TableView1 
      Height          =   4575
      Left            =   210
      TabIndex        =   1
      Top             =   1350
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   8070
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4575
      Left            =   4650
      TabIndex        =   0
      Top             =   1350
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   8070
      _Version        =   393217
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   10095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenDBFile 
         Caption         =   "Open MS-Database"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuExpand 
         Caption         =   "Expand Current View"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuLoadTable 
         Caption         =   "&Load Selected Table"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&View Table"
         Index           =   1
         WindowList      =   -1  'True
         Begin VB.Menu dbitem 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuSelField 
         Caption         =   "Show Selected Field"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutShow 
         Caption         =   "About DTVL"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Designed and developed by Chris Hatton. if you want to reuse this code please email me.
'chris@hatton.com

'make sure you have ADO 2.5 referenced

Option Explicit
Dim itmPrevious As Long 'parent child index.
Dim itmPreviousIndex As Long 'indexed item of parent item
Dim itmPreviousName As String 'name of parent item
Dim itmPreviousItem As String 'item selected
Dim CopyItem As String
Dim DelIndex As Long
Dim CutItem As Boolean

Private Sub dbitem_Click(Index As Integer)
Call ActiveListView(dbitem.Item(Index).Caption, FrmView.ListView1)
End Sub

Private Sub Form_Load()

Call PrepareDB

End Sub
Private Sub PrepareDB()
Call DBConnected

If DBConnected = True Then
    Label1 = "Connected to Local Database " & "(" & MSMDB & ")"
    Label2 = ""
    TreeView1.Nodes.Clear
    ProgressBar1.Value = 0
    Call GetLocalTables(TableView1)
    Call GetTableSatisitic(TableView1)

    
    
Else
    Label1 = "Not Connected to Database"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
Set cn = Nothing
End Sub

Private Sub mnuAboutShow_Click()
frmAbout.Show 1
End Sub

Private Sub mnuCopy_Click()
CopyItem = TreeView1.SelectedItem.Text
End Sub

Private Sub mnuCut_Click()
        Call HitTreeItem
        mnuCopy_Click
        CutItem = True
        TreeView1.SelectedItem.ForeColor = &H808080
        DelIndex = TreeView1.SelectedItem.Index
        TreeView1.Refresh

End Sub

Private Sub MnuExit_Click()
End
End Sub

Private Sub mnuExpand_Click()
TreeView1.SelectedItem.Expanded = True
End Sub

Private Sub mnuLoadTable_Click()
TableView1_DblClick
End Sub

Private Sub mnuOpenDBFile_Click()
On Error Resume Next
Comdiag.FileName = "*.mdb"
Comdiag.DialogTitle = "Microsoft Access Database"
Comdiag.CancelError = True
Comdiag.ShowOpen
If Err.Description = "Cancel was selected." Then Exit Sub

Custom = True
MSMDB = Comdiag.FileName

cn.Close
Set cn = Nothing
Call ClearMenuItems
Call PrepareDB
End Sub

Private Sub mnuPaste_Click()
On Error GoTo PasteErr
If Len(CopyItem) = 0 Then Exit Sub
TreeView1.SelectedItem.Expanded = True
Call AddSubValue(CopyItem)
        
        
If CutItem = True Then
    Call DelLocalValue("Select * from [" & itmPreviousItem & "]", itmPrevious - itmPreviousIndex, itmPreviousName)
    TreeView1.Nodes.Remove (DelIndex)
End If

CutItem = False
CopyItem = ""
Exit Sub
PasteErr:
MsgBox "Could not paste item"
End Sub

Private Sub mnuSelField_Click()
Call PopluateListView("Select " & TreeView1.SelectedItem.Text & " from [" & TableView1.SelectedItem.Text & "]", FrmView.ListView1)
FrmView.ListView1.ColumnHeaders.Item(1).Width = 5000
FrmView.Show
End Sub



Private Sub TableView1_DblClick()
Call ScopeRecords("Select * from [" & TableView1.SelectedItem.Text & "]", TreeView1)
End Sub

Private Sub TableView1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then Call ActiveListView(TableView1.SelectedItem.Text, FrmView.ListView1)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If TreeView1.Nodes.Count = 0 Then
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Else
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
End If
If TreeView1.SelectedItem.Parent.Selected = True Then Toolbar1.Buttons(3).Enabled = False Else Toolbar1.Buttons(3).Enabled = True


End Sub
Private Sub PrepareTreeUpdate()
     TreeView1.SetFocus
    ' TreeView1.SelectedItem.Expanded = True
  '   TreeView1.Nodes(TreeView1.SelectedItem.Index + 1).Selected = True
   '  If TreeView1.SelectedItem.LastSibling.Selected = False Then TreeView1.Nodes(TreeView1.SelectedItem.Index - 1).Selected = True
    ' If TreeView1.SelectedItem.FirstSibling.Index = 1 Then TreeView1.Nodes(TreeView1.SelectedItem.Index - 1).Selected = True

End Sub
Private Sub AddItem()
'  On Error GoTo AddUpdateErr
     Dim NewItem As String
     Call PrepareTreeUpdate 'organise the tree values first
     
     NewItem = InputBox("Type your Item", "Add New Item")
        
    
        If NewItem = "" Then
                Exit Sub
            Else
            
            Call HitTreeItem
            Call AddLocalValue("Select * from [" & itmPreviousItem & "]", NewItem, itmPreviousName)
            TreeView1.Nodes.Add TreeView1.SelectedItem.Parent.Index, tvwChild, , NewItem, 4
            TreeView1.Refresh

        End If
     Exit Sub
AddUpdateErr:
         Call AddSubValue(NewItem)
         

End Sub
Private Sub AddSubValue(NewString As String)
On Error GoTo SaveUpdateErr
  '      Call PrepareTreeUpdate
        Call HitTreeItem
        Call AddLocalSubValue("Select * from [" & itmPreviousItem & "]", NewString, itmPrevious - itmPreviousIndex, itmPreviousName, TreeView1)
Debug.Print itmPrevious
Exit Sub
SaveUpdateErr:
         MsgBox "Error Adding Current Value, Make sure your not adding a Auto Numbering type field", vbCritical + vbOKOnly
Exit Sub

End Sub

Private Sub DelItem()
On Error GoTo DelUpdateErr
 
    Dim DelQuest As String
    DelQuest = MsgBox("Delete Value?", vbCritical + vbOKCancel, "Delete Current Value")
    If DelQuest = vbOK Then
        Call HitTreeItem
        Call DelLocalValue("Select * from [" & itmPreviousItem & "]", itmPrevious - itmPreviousIndex, itmPreviousName)
        TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
    End If
    Exit Sub
DelUpdateErr:
    MsgBox "Error Deleting Current Value, Make Sure Your Not Deleting a Auto Numbering field or a Primary Key Field Type", vbCritical + vbOKOnly
    
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

Case 1
   Call AddItem
Case 2
  
    TreeView1.StartLabelEdit
Case 3
   Call DelItem
Case 5
    Call ActiveListView(TableView1.SelectedItem.Text, FrmView.ListView1)
End Select



End Sub
Private Sub ActiveListView(TableName As String, Lv As ListView)
Call PopluateListView("Select * from [" & TableName & "]", Lv)
FrmView.Show

End Sub
Private Sub HitTreeItem()
    On Error Resume Next
        itmPrevious = TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Previous.Index
        itmPreviousIndex = TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Parent.Index
        itmPreviousName = TreeView1.SelectedItem.Parent.Text
        itmPreviousItem = TableView1.SelectedItem.Text
If TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).FirstSibling.Selected = True Then itmPrevious = TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).FirstSibling.Index - 1
End Sub
Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo SaveUpdateErr
        Call HitTreeItem
        Call EditLocalValue("Select * from [" & itmPreviousItem & "]", NewString, itmPrevious - itmPreviousIndex, itmPreviousName)
Debug.Print itmPrevious
Exit Sub
SaveUpdateErr:
Cancel = 1
MsgBox "Could not Save Changes, Check your Database field Type", vbCritical + vbOKOnly

End Sub

Private Sub TreeView1_DblClick()

'If TreeView1.SelectedItem.FirstSibling.Selected = True Or TreeView1.SelectedItem.LastSibling.Selected = True Then TreeView1.StartLabelEdit
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case 46
    Call DelItem
    
Case 112
    Call AddItem
    
Case 113
    TreeView1.StartLabelEdit
    
Case 114
    Call DelItem
Case 115
    Call ActiveListView(TableView1.SelectedItem.Text, FrmView.ListView1)

End Select
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuView
End Sub
