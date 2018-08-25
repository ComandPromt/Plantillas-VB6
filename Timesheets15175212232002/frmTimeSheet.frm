VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimeSheet 
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9435
   ControlBox      =   0   'False
   Icon            =   "frmTimeSheet.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   9435
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboWeek 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddd, dd MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   24
      ToolTipText     =   "Week beginning for current timesheet"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      ToolTipText     =   "Click to accept changes"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      ToolTipText     =   "Click to cancel changes since last OK"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      ToolTipText     =   "Click to exit the form"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Day"
      Height          =   615
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Click to select a day of the week"
      Top             =   480
      Width           =   7215
      Begin VB.OptionButton optSunday 
         Caption         =   "Sunday"
         Height          =   255
         Left            =   6240
         TabIndex        =   6
         ToolTipText     =   "Click to enter data for Sunday"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSaturday 
         Caption         =   "Saturday"
         Height          =   255
         Left            =   5280
         TabIndex        =   5
         ToolTipText     =   "Click to enter data for Saturday"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFriday 
         Caption         =   "Friday"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         ToolTipText     =   "Click to enter data forFriday"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptThursday 
         Caption         =   "Thursday"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         ToolTipText     =   "Click to enter data for Thursday"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optWednesday 
         Caption         =   "Wednesday"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Click to enter data for Wednesday"
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTuesday 
         Caption         =   "Tuesday"
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Click to enter data for Tuesday"
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optMonday 
         Caption         =   "Monday"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Click to enter data for Monday"
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame fraDay 
      Caption         =   "Monday"
      Height          =   3735
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   7215
      Begin MSComctlLib.ListView lvwTimes 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Times for the day."
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3625
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Project"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time (hrs)"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtTime 
         Height          =   285
         Left            =   3240
         TabIndex        =   12
         Text            =   "3.5"
         ToolTipText     =   "Time spent (hours)"
         Top             =   3345
         Width           =   495
      End
      Begin VB.CommandButton btnDeleteEntry 
         Caption         =   "Delete Entry"
         Height          =   375
         Left            =   180
         TabIndex        =   9
         ToolTipText     =   "Click to delete an entry"
         Top             =   3090
         Width           =   1095
      End
      Begin VB.CommandButton btnAddEntry 
         Caption         =   "Add Entry"
         Height          =   375
         Left            =   180
         TabIndex        =   8
         ToolTipText     =   "Click to add an entry"
         Top             =   2610
         Width           =   1095
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   3240
         TabIndex        =   11
         Text            =   "Admin"
         ToolTipText     =   "Description of time spent"
         Top             =   3000
         Width           =   3765
      End
      Begin VB.ComboBox cboProject 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Project Number"
         Top             =   2640
         Width           =   3765
      End
      Begin VB.Label Label7 
         Caption         =   "Project"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblTotal 
         Caption         =   "8 hrs"
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "hrs"
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Time"
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   3000
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Week Beginning"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   150
      Width           =   1335
   End
End
Attribute VB_Name = "frmTimeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddEntry_Click()
  intFormAction = ADD_NEW
  timesheetCode.enableTimeDisplay
End Sub

Private Sub btnCancel_Click()
  timOld.copyTime timCurrent
  timCurrent.display
End Sub

Private Sub btnDeleteEntry_Click()
  Dim intReturn As Integer
  If timCurrent.lngID <= 0 Then Exit Sub
  intReturn = MsgBox("Are you sure you want to delete this entry?", vbYesNo)
  If intReturn = vbNo Then Exit Sub
  timCurrent.deleteTime timCurrent.lngID
  timesheetCode.fillTimesListview Me.lvwTimes, Me.cboWeek
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
  intFormAction = 0
End Sub

Private Sub btnExit_Click()
  Unload Me
End Sub

Private Sub btnOK_Click()
  Dim itmTmp As ListItem
  Select Case intFormAction
    Case ADD_NEW
      If Len(timCurrent.strDescription) = 0 Or timCurrent.dblHours = 0 Or Me.cboProject.ListIndex = -1 Then
        MsgBox "Not all fields filled in correctly, return and correct."
        Exit Sub
      End If
      timCurrent.addTime timesheetCode.returnProjectID(Me.cboProject)
    Case EDIT
      If timCurrent.blnInvoiced = True Then
        MsgBox "You cannot edit this time as it has already been invoiced"
        Exit Sub
      End If
      Set itmTmp = lvwTimes.ListItems(checkSelected(Me.lvwTimes))
      timCurrent.editTime CLng(Right(itmTmp.Key, Len(itmTmp.Key) - Len("Item_")))
    Case DELETE
  End Select
  timesheetCode.fillTimesListview Me.lvwTimes, Me.cboWeek
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
  intFormAction = 0
End Sub

Private Sub cboProject_Change()
  timCurrent.lngProjectID = timesheetCode.returnProjectID(Me.cboProject)
End Sub

Private Sub cboProject_Click()
  timCurrent.lngProjectID = timesheetCode.returnProjectID(Me.cboProject)
End Sub

Private Sub cboWeek_Change()
  timCurrent.clear
  timOld.clear
  timesheetCode.fillTimesListview Me.lvwTimes, Me.cboWeek
  timesheetCode.fillProjectCombo Me.cboProject
  timesheetCode.disableTimeDisplay
End Sub

Private Sub cboWeek_Click()
  timCurrent.clear
  timOld.clear
  timesheetCode.fillTimesListview Me.lvwTimes, Me.cboWeek
  timesheetCode.fillProjectCombo Me.cboProject
  timesheetCode.disableTimeDisplay
End Sub

Private Sub Form_Load()
  fillDateCombo Me.cboWeek
  Me.cboWeek.ListIndex = 0
  intTimesheetWeekday = vbMonday
  fraDay.Caption = "Monday"
  timCurrent.clear
  timOld.clear
  timesheetCode.fillTimesListview Me.lvwTimes, Me.cboWeek
  timesheetCode.fillProjectCombo Me.cboProject
  timesheetCode.disableTimeDisplay
  
  Select Case Weekday(Date)
    Case vbMonday
      Me.optMonday.Value = True
    Case vbTuesday
      Me.optTuesday.Value = True
    Case vbWednesday
      Me.optWednesday.Value = True
    Case vbThursday
      Me.OptThursday.Value = True
    Case vbFriday
      Me.optFriday.Value = True
    Case vbSaturday
      Me.optSaturday.Value = True
    Case vbSunday
      Me.optSunday.Value = True
  End Select
End Sub

Private Sub lvwTimes_Click()
  Dim itmTmp As ListItem
  If checkSelected(Me.lvwTimes) <> -1 Then
    timesheetCode.enableTimeDisplay
    Set itmTmp = lvwTimes.ListItems(checkSelected(Me.lvwTimes))
    intFormAction = EDIT
    timCurrent.loadTime CLng(Right(itmTmp.Key, Len(itmTmp.Key) - Len("Item_")))
    timCurrent.display
    timCurrent.copyTime timOld
  End If
End Sub

Private Sub optFriday_Click()
  intTimesheetWeekday = vbFriday
  fillTimesListview Me.lvwTimes, Me.cboWeek
  fraDay.Caption = "Friday " & (datSelectedDate)
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
End Sub

Private Sub optMonday_Click()
  intTimesheetWeekday = vbMonday
  fillTimesListview Me.lvwTimes, Me.cboWeek
  fraDay.Caption = "Monday " & (datSelectedDate)
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
End Sub

Private Sub optSaturday_Click()
  intTimesheetWeekday = vbSaturday
  fillTimesListview Me.lvwTimes, Me.cboWeek
  fraDay.Caption = "Saturday " & (datSelectedDate)
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
End Sub

Private Sub optSunday_Click()
  intTimesheetWeekday = vbSunday
  fillTimesListview Me.lvwTimes, Me.cboWeek
  fraDay.Caption = "Sunday " & (datSelectedDate)
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
End Sub

Private Sub OptThursday_Click()
  intTimesheetWeekday = vbThursday
  fillTimesListview Me.lvwTimes, Me.cboWeek
  fraDay.Caption = "Thursday " & (datSelectedDate)
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
End Sub

Private Sub optTuesday_Click()
  intTimesheetWeekday = vbTuesday
  fillTimesListview Me.lvwTimes, Me.cboWeek
  fraDay.Caption = "Tuesday " & (datSelectedDate)
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
End Sub

Private Sub optWednesday_Click()
  intTimesheetWeekday = vbWednesday
  fillTimesListview Me.lvwTimes, Me.cboWeek
  fraDay.Caption = "Wednesday " & (datSelectedDate)
  timCurrent.clear
  timOld.clear
  timCurrent.display
  timesheetCode.disableTimeDisplay
End Sub

Private Sub txtDescription_Change()
  If Len(Me.txtDescription) > 0 Then timCurrent.strDescription = Me.txtDescription
End Sub

Private Sub txtDescription_Validate(Cancel As Boolean)
  If Len(Me.txtDescription) < 1 Then
    MsgBox "Description Required"
    Cancel = True
  End If
End Sub

Private Sub txtTime_Change()
  If IsNumeric(Me.txtTime) = True Then
    timCurrent.dblHours = Me.txtTime
  End If
End Sub

Private Sub txtTime_Validate(Cancel As Boolean)
  If IsNumeric(Me.txtTime) = False Then
    MsgBox "Numeric Value Required"
    Me.txtTime = timCurrent.dblHours
    Cancel = True
  End If
End Sub
