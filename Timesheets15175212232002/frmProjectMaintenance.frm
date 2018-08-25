VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjectMaintenance 
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9435
   ControlBox      =   0   'False
   Icon            =   "frmProjectMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   9435
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCustomer 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Text            =   "Text1"
      ToolTipText     =   "The customer, leave blank for internal jobs"
      Top             =   1800
      Width           =   2655
   End
   Begin MSComctlLib.ListView lvwProjects 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "A list of all the projects in the system, click to select one"
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Number"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2002
      EndProperty
   End
   Begin VB.Frame fraFinancial 
      Caption         =   "Financial Budget"
      Height          =   1815
      Left            =   6120
      TabIndex        =   24
      Top             =   2880
      Width           =   3255
      Begin VB.TextBox txtEstimatedTravel 
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Text            =   "$500"
         ToolTipText     =   "Estimated Travel Costs"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtEstimatedMaterial 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Text            =   "$6000"
         ToolTipText     =   "Esitmated Material Costs"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtEstimatedLabour 
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Text            =   "$21,400"
         ToolTipText     =   "Estimated labour costs"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Label12"
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1455
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3120
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label10 
         Caption         =   "Esitmated Travel Costs"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   975
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Estimated Material Costs"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   615
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Estimated Labour Costs"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   255
         Width           =   1935
      End
   End
   Begin VB.TextBox txtProjectNumber 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "P110201"
      ToolTipText     =   "The Project Number (required)"
      Top             =   105
      Width           =   1935
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Click to delete a project"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "New"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click to add a new project"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      ToolTipText     =   "Click to exit the form"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      ToolTipText     =   "Click to cancel the changes since the last OK"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      ToolTipText     =   "Click to accept the changes"
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cboManager 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Who the manager is"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.ComboBox cboCreatedBy 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Who created this project"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtDateClosed 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      ToolTipText     =   "Date project was closed"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtDateCreated 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Text            =   "21/11/2001"
      ToolTipText     =   "Date the project was created"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtProjectDescription 
      Height          =   885
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmProjectMaintenance.frx":000C
      ToolTipText     =   "The project description"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtProjectName 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Text            =   "Test Project"
      ToolTipText     =   "The project name (required)"
      Top             =   465
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Customer"
      Height          =   255
      Left            =   2160
      TabIndex        =   30
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Project Number"
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Project Manager"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Created By"
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   2190
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Date Closed"
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   3255
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Date Created"
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   2895
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Project Description"
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Project Name"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmProjectMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private curRunningTotal As Currency

Private Sub btnCancel_Click()
  proOld.copyProject proCurrent
  proCurrent.display
End Sub

Private Sub btnDelete_Click()
  Dim strMessage As String
  If checkSelected(Me.lvwProjects) <> -1 Then
    strMessage = "Warning: all times associated with this project will also be deleted" & Chr(10) & Chr(10) & "Are you sure you wish to continue?"
    If MsgBox(strMessage, vbYesNo, "Deletion Warning") = vbYes Then
      timesheetCode.deleteProjectTimes getIDFromKey(Me.lvwProjects.SelectedItem.Key)
      proCurrent.deleteProject
      globalCode.fillProjectListView Me.lvwProjects
      proCurrent.clear
      proOld.clear
      proCurrent.display
      projectMaintenanceCode.disableDisplay
    End If
  End If
End Sub

Private Sub btnExit_Click()
  Unload Me
End Sub

Private Sub btnNew_Click()
  intFormAction = ADD_NEW
  proCurrent.clear
  proOld.clear
  proCurrent.display
  projectMaintenanceCode.enableDisplay
  Me.txtDateCreated = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
End Sub

Private Sub btnOK_Click()
  Dim lngSelectedItem As Long
  lngSelectedItem = 0
  Select Case intFormAction
    Case ADD_NEW
      If Len(proCurrent.strProjectNumber) = 0 Or Len(proCurrent.strProjectName) = 0 Then
        MsgBox "Not all fields filled in correctly, please return and correct"
        Exit Sub
      End If
      proCurrent.addProject
      lngSelectedItem = proCurrent.getLatestProject
    Case EDIT
      proCurrent.editProject getIDFromKey(Me.lvwProjects.SelectedItem.Key)
      lngSelectedItem = getIDFromKey(Me.lvwProjects.SelectedItem.Key)
    Case DELETE
  End Select
  globalCode.fillProjectListView Me.lvwProjects, lngSelectedItem
  proCurrent.display
  intFormAction = EDIT
End Sub

Private Sub cboCreatedBy_Change()
  proCurrent.lngCreatedByID = projectMaintenanceCode.returnUserID(Me.cboCreatedBy)
  proCurrent.strCreatedByName = Me.cboCreatedBy
End Sub

Private Sub cboCreatedBy_Click()
  proCurrent.lngCreatedByID = projectMaintenanceCode.returnUserID(Me.cboCreatedBy)
  proCurrent.strCreatedByName = Me.cboCreatedBy

End Sub

Private Sub cboManager_Change()
  proCurrent.lngManagerID = projectMaintenanceCode.returnUserID(Me.cboManager)
  proCurrent.strManagerName = Me.cboManager
End Sub

Private Sub cboManager_Click()
  proCurrent.lngManagerID = projectMaintenanceCode.returnUserID(Me.cboManager)
  proCurrent.strManagerName = Me.cboManager

End Sub

Private Sub Form_Load()
  globalCode.fillProjectListView Me.lvwProjects
  proCurrent.clear
  proOld.clear
  proCurrent.display
  projectMaintenanceCode.fillUserCombo Me.cboCreatedBy
  projectMaintenanceCode.fillUserCombo Me.cboManager
  projectMaintenanceCode.disableDisplay
  intFormAction = 0
  If checkSelected(Me.lvwProjects) <> -1 Then Me.lvwProjects.ListItems(checkSelected(Me.lvwProjects)).Selected = False
End Sub

Private Sub lvwProjects_Click()
  If checkSelected(Me.lvwProjects) <> -1 Then
    projectMaintenanceCode.enableDisplay
    intFormAction = EDIT
    proCurrent.loadProject getIDFromKey(Me.lvwProjects.SelectedItem.Key)
    proCurrent.display
    proCurrent.copyProject proOld
  End If
End Sub


Private Sub txtCustomer_Change()
  proCurrent.strCustomer = Me.txtCustomer
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
  If Len(Me.txtCustomer) > 50 Then
    MsgBox "Valid Customer Name Required, length must be < 51 characters"
    Cancel = True
  End If
End Sub

Private Sub txtDateClosed_Change()
  If IsDate(Me.txtDateClosed) = True Then
    proCurrent.datClosed = Me.txtDateClosed
  End If

End Sub

Private Sub txtDateClosed_Validate(Cancel As Boolean)
  If IsDate(Me.txtDateClosed) = False Then
    MsgBox "Valid Date Required"
    Cancel = True
  End If
End Sub

Private Sub txtDateCreated_Change()
  If IsDate(Me.txtDateCreated) = True Then
    proCurrent.datCreated = Me.txtDateCreated
  End If
End Sub

Private Sub txtDateCreated_Validate(Cancel As Boolean)
  If IsDate(Me.txtDateCreated) = False Then
    MsgBox "Valid Date Required"
    Cancel = True
  End If
End Sub

Private Sub txtEstimatedLabour_Change()
  If IsNumeric(Me.txtEstimatedLabour) = True Then
    proCurrent.curBudgetLabour = Me.txtEstimatedLabour
    curRunningTotal = CCur(Me.txtEstimatedLabour) + CCur(Me.txtEstimatedMaterial) + CCur(Me.txtEstimatedTravel)
    Me.lblTotal = curRunningTotal
  Else
    curRunningTotal = CCur(Me.txtEstimatedMaterial) + CCur(Me.txtEstimatedTravel)
    Me.lblTotal = curRunningTotal
  End If
End Sub

Private Sub txtEstimatedLabour_Validate(Cancel As Boolean)
  If IsNumeric(Me.txtEstimatedLabour) = False Then
    MsgBox "Numeric Value Required"
    Me.txtEstimatedLabour = proCurrent.curBudgetLabour
    Cancel = True
  End If
End Sub

Private Sub txtEstimatedMaterial_Change()
  If IsNumeric(Me.txtEstimatedMaterial) = True Then
    proCurrent.curBudgetMaterial = Me.txtEstimatedMaterial
    curRunningTotal = CCur(Me.txtEstimatedLabour) + CCur(Me.txtEstimatedMaterial) + CCur(Me.txtEstimatedTravel)
    Me.lblTotal = curRunningTotal
  Else
    curRunningTotal = CCur(Me.txtEstimatedLabour) + CCur(Me.txtEstimatedTravel)
    Me.lblTotal = curRunningTotal
  End If
End Sub

Private Sub txtEstimatedMaterial_Validate(Cancel As Boolean)
  If IsNumeric(Me.txtEstimatedMaterial) = False Then
    MsgBox "Numeric Value Required"
    Me.txtEstimatedMaterial = proCurrent.curBudgetMaterial
    Cancel = True
  End If

End Sub

Private Sub txtEstimatedTravel_Change()
  If IsNumeric(Me.txtEstimatedTravel) = True Then
    proCurrent.curBudgetTravel = Me.txtEstimatedTravel
    curRunningTotal = CCur(Me.txtEstimatedLabour) + CCur(Me.txtEstimatedMaterial) + CCur(Me.txtEstimatedTravel)
    Me.lblTotal = curRunningTotal
  Else
    curRunningTotal = CCur(Me.txtEstimatedLabour) + CCur(Me.txtEstimatedMaterial)
    Me.lblTotal = curRunningTotal
  End If
End Sub

Private Sub txtEstimatedTravel_Validate(Cancel As Boolean)
  If IsNumeric(Me.txtEstimatedTravel) = False Then
    MsgBox "Numeric Value Required"
    Me.txtEstimatedTravel = proCurrent.curBudgetTravel
    Cancel = True
  End If
End Sub

Private Sub txtProjectDescription_Change()
  proCurrent.strProjectDescription = Me.txtProjectDescription
End Sub

Private Sub txtProjectName_Change()
  proCurrent.strProjectName = Me.txtProjectName
End Sub

Private Sub txtProjectName_Validate(Cancel As Boolean)
  If Len(Me.txtProjectName) < 1 Or Len(Me.txtProjectName) > 50 Then
    MsgBox "Valid Project Name Required, length must be > 1 character and < 51 characters"
    Cancel = True
  End If
End Sub

Private Sub txtProjectNumber_Change()
  proCurrent.strProjectNumber = Me.txtProjectNumber
End Sub

Private Sub txtProjectNumber_Validate(Cancel As Boolean)
  If Len(Me.txtProjectNumber) < 1 Or Len(Me.txtProjectNumber) > 10 Then
    MsgBox "Valid Project Number Required, length must be > 1 character and < 11 characters"
    Cancel = True
    Me.txtProjectNumber = proCurrent.strProjectNumber
  End If
End Sub
