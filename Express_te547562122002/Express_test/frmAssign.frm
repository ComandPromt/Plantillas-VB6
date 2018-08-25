VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmAssign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmAssign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "1"
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "1"
      Top             =   3840
      Width           =   375
   End
   Begin VB.Data dat_ass 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Assign Test"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "frmAssign.frx":0442
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2514
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test Expires after"
      Height          =   1095
      Left            =   600
      TabIndex        =   7
      Top             =   3360
      Width           =   3375
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Days"
         Height          =   195
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Months"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   525
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Student Name"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   1920
      Picture         =   "frmAssign.frx":0456
      Top             =   120
      Width           =   2700
   End
End
Attribute VB_Name = "frmAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub load_list()
With dat_ass
If .Recordset.AbsolutePosition < 0 Then
        'MsgBox "This Question Bank is Empty"
        Exit Sub
        End If
    ''''''''''''''''''''''''''''''''''''
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
    If .Recordset.Fields("test_name") = frmtest_creator.test_name Then
    List1.AddItem (.Recordset.Fields("UserID"))
    End If
    .Recordset.MoveNext
    Loop
    '''''''''''''''''''''''''''''''''''''
 End With
 
End Sub
Private Sub Command1_Click()
If DBList1.Text = "" Then
Exit Sub
Else
With dat_ass.Recordset
.AddNew
.Fields("UserID") = DBList1.Text
.Fields("test_name") = frmtest_creator.test_name
.Fields("file_path") = frmtest_creator.bank_filename
.Fields("e_date") = DateAdd("m", Val(Text1), DateAdd("d", Val(Text2), Date))

End With
On Error GoTo err_dup
dat_ass.UpdateRecord
List1.AddItem DBList1.Text

err_dup:
If Err = 3022 Then
MsgBox "This Test is already Assigned to the Student", vbOKOnly, "Warning !!"
dat_ass.Recordset.CancelUpdate
'List1.Clear
Exit Sub
End If
Resume Next




End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

    dat_ass.DatabaseName = App.Path & "\login.exp"
    dat_ass.RecordSource = "Test_Assign"
    dat_ass.Refresh
    load_list
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
frmAssign.Caption = "Assign " & frmtest_creator.test_name
    Data1.DatabaseName = App.Path & "\login.exp"
    Data1.RecordSource = "select UserID from Login where Instructor=False"
    Data1.Refresh
    DBList1.ListField = "UserID"
    
  
End Sub
