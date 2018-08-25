VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmreport_t 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Report"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   Icon            =   "frmreport_t.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
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
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancle"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "frmreport_t.frx":0442
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3889
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select a Test"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmreport_t"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim headder As String
Dim table_body As String
Dim marquee As String
Private Sub Command1_Click()
Dim dblist_txt As String
Dim no_test As Integer
Dim total_p As Integer
dblist_txt = DBList1.Text

table_body = "<center><table BORDER=5 COLS=4 WIDTH=""100%"" ><tr><td>Student Name</td><td>Attempt</td><td>Score %</td><td>Attempt Date</td></tr>"
headder = "<html><head><title>Test Report</title></head>" _
& "<center><b><u><font size=+1>TEST REPORT</font></u></b>" _
& "<br>Test Name :" & DBList1.Text & "<br>Date :" & Date & "<br>Time :" & Time & "<br><hr WIDTH=""100%""></center><br>&nbsp;"
With Data1
    .DatabaseName = App.Path & "\login.exp"
    .RecordSource = "Test_Assign"
    .Refresh
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
    If .Recordset.Fields("test_name") = dblist_txt Then
'''''''''''''''''''''''''''''''''''''
table_body = table_body & "<tr><td>" & .Recordset.Fields("UserID") & "</td>"
     If .Recordset.Fields("attempt") Then
     no_test = no_test + 1
     total_p = total_p + .Recordset.Fields("score_p")

table_body = table_body & "<td>YES</td>" & "<td>" & .Recordset.Fields("score_p") & "</td>" & "<td>" & .Recordset.Fields("a_date") & "</td>"

     Else
table_body = table_body & "<td>NO</td>" & "<td>-</td>" & "<td>-</td>"
     End If

'  '''''''''''''   MsgBox .Recordset.Fields("file_path")
    End If
    .Recordset.MoveNext
    Loop

End With

marquee = "<p align=""center""><marquee width=""50%"">Average Score %=" & (total_p / no_test) & "</marquee><br>&nbsp; </p>"
Open App.Path & "\" & dblist_txt & ".htm" For Output As #3
Print #3, headder
Print #3, marquee
Print #3, table_body
Close 3
ShellExecute Me.hwnd, vbNullString, App.Path & "\" & dblist_txt & ".htm", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
With Data2
        .DatabaseName = frmtest_creator.bank_filename
        .RecordSource = "all_test"
        .Refresh
        End With
DBList1.ListField = "test_name"
End Sub


