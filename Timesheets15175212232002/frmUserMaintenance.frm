VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserMaintenance 
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9435
   ControlBox      =   0   'False
   Icon            =   "frmUserMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   9435
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboSecurityLevel 
      Height          =   315
      ItemData        =   "frmUserMaintenance.frx":000C
      Left            =   4320
      List            =   "frmUserMaintenance.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Click to select the security level for this user"
      Top             =   1200
      Width           =   3885
   End
   Begin MSComctlLib.ListView lvwUsers 
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "A list of all the users in the system, click to select one."
      Top             =   90
      Width           =   2535
      _ExtentX        =   4471
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
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Abb."
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.TextBox txtAbbreviation 
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Text            =   "PVL"
      ToolTipText     =   "The users abbreviation (required)"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Text            =   "van de Loo"
      ToolTipText     =   "The users last name (required)"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Text            =   "Paul"
      ToolTipText     =   "The users first name (required)"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      ToolTipText     =   "Click to accept the changes"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      ToolTipText     =   "Click to cancel the changes since the last OK"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      ToolTipText     =   "Click to exit the form"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "New"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click to add a new user"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Click to delete a user"
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Security Level"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   13
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Abbreviation"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmUserMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
  usrOld.copyUser usrCurrent
  usrCurrent.display
End Sub

Private Sub btnDelete_Click()
  Dim strMessage As String
  If checkSelected(Me.lvwUsers) <> -1 Then
    strMessage = "Warning: all times associated with this user will also be deleted" & Chr(10) & Chr(10) & "Are you sure you wish to continue?"
    If MsgBox(strMessage, vbYesNo, "Deletion Warning") = vbYes Then
      timesheetCode.deleteUserTimes getIDFromKey(Me.lvwUsers.SelectedItem.Key)
      usrCurrent.deleteUser
      globalCode.fillUserListView Me.lvwUsers
      usrCurrent.clear
      usrOld.clear
      usrCurrent.display
      userMaintenanceCode.disableDisplay
    End If
  End If
End Sub

Private Sub btnExit_Click()
  Unload Me
End Sub

Private Sub btnNew_Click()
  intFormAction = ADD_NEW
  usrCurrent.clear
  usrOld.clear
  usrCurrent.display
  userMaintenanceCode.enableDisplay
End Sub

Private Sub btnOK_Click()

  Select Case intFormAction
    Case ADD_NEW
      If Len(usrCurrent.strFirstName) = 0 Or Len(usrCurrent.strLastName) = 0 Or Len(usrCurrent.strAbbreviation) = 0 Then
        MsgBox "Not all fields filled in correctly, return and correct."
        Exit Sub
      End If
      usrCurrent.addUser
    Case EDIT
      usrCurrent.editUser globalCode.getIDFromKey(Me.lvwUsers.SelectedItem.Key)
    Case DELETE
  End Select
  globalCode.fillUserListView Me.lvwUsers
  usrCurrent.clear
  usrOld.clear
  usrCurrent.display
  userMaintenanceCode.disableDisplay
  intFormAction = 0

End Sub

Private Sub cboSecurityLevel_click()
  usrCurrent.bytSecurityLevel = Me.cboSecurityLevel.ListIndex + 1
End Sub

Private Sub Form_Load()
  globalCode.fillUserListView Me.lvwUsers
  usrCurrent.clear
  usrOld.clear
  usrCurrent.display
  userMaintenanceCode.disableDisplay
  intFormAction = 0
End Sub

Private Sub lvwUsers_Click()
  If checkSelected(Me.lvwUsers) <> -1 Then
    userMaintenanceCode.enableDisplay
    intFormAction = EDIT
    usrCurrent.loadUser globalCode.getIDFromKey(Me.lvwUsers.SelectedItem.Key)
    usrCurrent.display
    usrCurrent.copyUser usrOld
  End If
End Sub

Private Sub txtAbbreviation_Change()
  If Len(Me.txtFirstName) > 0 Then usrCurrent.strAbbreviation = Me.txtAbbreviation
End Sub

Private Sub txtAbbreviation_Validate(Cancel As Boolean)
  If Len(Me.txtAbbreviation) < 1 Then
    MsgBox "Abbreviation Required"
    Cancel = True
  End If
End Sub

Private Sub txtFirstName_Change()
  If Len(Me.txtFirstName) > 0 Then usrCurrent.strFirstName = Me.txtFirstName
End Sub

Private Sub txtFirstName_Validate(Cancel As Boolean)
  If Len(Me.txtFirstName) < 1 Or Len(Me.txtFirstName) > 30 Then
    MsgBox "Valid First Name Required, length must be > 1 character and < 31 characters"
    Cancel = True
  End If
End Sub


Private Sub txtLastName_Change()
  If Len(Me.txtLastName) > 0 Then usrCurrent.strLastName = Me.txtLastName
End Sub

Private Sub txtLastName_Validate(Cancel As Boolean)
  If Len(Me.txtLastName) < 1 Or Len(Me.txtLastName) > 30 Then
    MsgBox "Valid Last Name Required, length must be > 1 character and < 31 characters"
    Cancel = True
  End If
End Sub
