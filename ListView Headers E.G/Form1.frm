VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Inputbox Code
Dim name 'Remember name
Dim age 'Remember age
Dim dob 'remember dob
name = InputBox("What is the persons name ?", "Name ?") 'Show a inputbox for the persons name
age = InputBox("What is the personsage ?", "Age ?") 'Show a inputbox for the persons age
dob = InputBox("What is the persons Date Of Birth ?", "Date Of Birth ?") 'Show a inputbox for the persons Date Of Birth
'End Of Inputbox Code

'Adding To List Code
Dim ListObj As ListItem 'Set listObj as a listitem
Set ListObj = ListView1.ListItems.Add(, , name) 'this allways adds to the 1st column , this lines adds the name to the 1st Column
ListObj.SubItems(1) = age 'this allways adds to the second column , this lines adds the age to the 2nd Column
ListObj.SubItems(2) = dob 'this allways adds to the third column , this lines adds the dob to theb 3rd Column
'End Of Adding To List Code

End Sub

Private Sub Command2_Click()
On Error Resume Next 'If Theres a error resume the next line ( the error here would be nothing in the listview or no selected item )
ListView1.ListItems.Remove ListView1.SelectedItem.Index 'Delete The Selected Item
End Sub

Private Sub Command3_Click()
Unload Me 'Exits The Program
End Sub

Private Sub Form_Load()
'This Code is needed
ListView1.View = lvwReport 'Set The Listview1 View So we Can See Our Columns/Headers
ListView1.ColumnHeaders.Add , , "Name" 'Add a column Called Name
ListView1.ColumnHeaders.Add , , "Age" 'Add a column Called Age
ListView1.ColumnHeaders.Add , , "Date Of Birth" 'Add a column Called Age
'End Of Needed Code
End Sub

Private Sub Form_Unload(Cancel As Integer)
'My Code
If MsgBox("If This Code Helped You Please Come Back and Either Vote Or Comment,:). Would You Like To Vote Or Comment Now?", vbYesNo, "Thankx For Using My Code") = vbYes Then GoTo open_url
Exit Sub
open_url:
MsgBox "I have Copyed The Url To The clipboard", vbInformation, "Url"
Clipboard.SetText "http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=36543&lngWId=1"
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next ' resume The Next line On a Error
MsgBox "Name : " + ListView1.SelectedItem.Text + vbCrLf + "Age : " + ListView1.SelectedItem.SubItems(1) + vbCrLf + "Dob : " + ListView1.SelectedItem.SubItems(2) 'Make The Msgbox
End Sub
