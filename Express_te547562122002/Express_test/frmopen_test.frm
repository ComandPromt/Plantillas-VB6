VERSION 5.00
Begin VB.Form frmopen_test 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Test"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmopen_test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   " Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Data datall 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Open"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Select Test"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmopen_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.Text = "" Then
MsgBox "Please select a Test", , "Nothing to Open!"
Else
frmtest_creator.load_test List1.Text
Unload Me
End If
End Sub

Private Sub Command2_Click()
If List1.Text = "" Then
MsgBox "Please select a Test", , "Nothing to Delete!"
Else
''''''''''''''''''''''''''''''''''''''''''''
'This function will show how to delete records
Dim DB 'As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field
'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(frmtest_creator.bank_filename)
'Set the table to open
'On Error Resume Next
DB.TableDefs.Delete List1.Text
'close the database
DB.Close
''''''''''''''''''''''''''''''''''''''''''''

With datall
.Recordset.FindFirst ("test_name='" & List1.Text & "'")
.Recordset.Delete
.UpdateRecord
End With

List1.Clear
Form_Load
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
With datall
        .DatabaseName = frmtest_creator.bank_filename
        .RecordSource = "all_test"
        .Refresh
        If .Recordset.AbsolutePosition < 0 Then Exit Sub
    ''''''''''''''''''''''''''''''''''''
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
    List1.AddItem (.Recordset.Fields("test_name"))
    .Recordset.MoveNext
    Loop
    '''''''''''''''''''''''''''''''''''''''
        
        
    End With
End Sub
