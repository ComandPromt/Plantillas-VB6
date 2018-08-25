VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmreport_i 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instructor Report"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "frmreport_i.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
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
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancle"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "frmreport_i.frx":0442
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3545
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select an Instructor"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "frmreport_i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim headder As String
Dim table_body As String
Dim marquee As String
Private Sub Command1_Click()
Dim no_s As Integer
Dim no_test As Integer
Dim total_p As Integer
Dim test As String
table_body = "<center><table BORDER=5 COLS=4 WIDTH=""100%"" ><tr><td>Test Name</td><td>Date</td><td>No Students</td><td>Avg Score</td></tr>"

headder = "<html><head><title>Instructor Report</title></head>" _
& "<center><b><u><font size=+1>INSTRUCTOR REPORT</font></u></b>" _
& "<br>UserID :" & DBList1.Text & "<br>Date :" & Date & "<br>Time :" & Time & "<br><hr WIDTH=""100%""></center><br>&nbsp;"
''''''''''''''''''''''''''''''''''''
With Data2
        .DatabaseName = frmtest_creator.bank_filename
        .RecordSource = "select * from all_test where User_ID='" & DBList1.Text & "'"
        .Refresh
End With
With Data3
        .DatabaseName = App.Path & "\login.exp"
        .RecordSource = "Test_Assign"
        .Refresh
End With
 
 Do Until Data2.Recordset.EOF
test = Data2.Recordset.Fields("test_name")
table_body = table_body & "<tr><td>" & test & "</td>" & "<td>" & Data2.Recordset.Fields("date") & "</td>"

no_s = 0
no_test = 0
total_p = 0
'''''''''''''''''''''''''''inner loop
Do Until Data3.Recordset.EOF
    If Data3.Recordset.Fields("test_name") = test Then
        If Data3.Recordset.Fields("attempt") = True Then
        no_s = no_s + 1
        no_test = no_test + 1
        total_p = total_p + Data3.Recordset.Fields("score_p")
        End If
     End If
    Data3.Recordset.MoveNext
    Loop 'inner loop
    If no_test = 0 Then no_test = 1
table_body = table_body & "<td>" & no_test & "</td>" & "<td>" & total_p / no_test & "</td>"

Data2.Recordset.MoveNext
Loop 'end of big loop


Open App.Path & "\" & DBList1.Text & ".htm" For Output As #2
Print #2, headder
'Print #2, marquee
Print #2, table_body
Close 2
'ShellExecute Me.hwnd, vbNullString, App.Path & "\" & DBList1.Text & ".htm", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
With Data1
        .DatabaseName = App.Path & "\login.exp"
        .RecordSource = "SELECT UserID, Password, FirstName, LastName " & _
                                "FROM Login " & _
                              "WHERE Instructor = True"
        .Refresh
        End With
DBList1.ListField = "UserID"
End Sub
