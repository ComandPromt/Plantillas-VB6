VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmreport_s 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Progress Report"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   Icon            =   "frmreport_s.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancle"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "frmreport_s.frx":0442
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   3545
      _Version        =   393216
   End
   Begin VB.Data datlogin 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select a Student"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmreport_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim headder As String
Dim table_body As String
Dim marquee As String
Private Sub Command1_Click()
Dim no_test As Integer
Dim total_p As Integer

table_body = "<center><table BORDER=5 COLS=4 WIDTH=""100%"" ><tr><td>Test Name</td><td>Attempt</td><td>Score %</td><td>Expire Date</td></tr>"

headder = "<html><head><title>Student Progress Report</title></head>" _
& "<center><b><u><font size=+1>STUDENT REPORT</font></u></b>" _
& "<br>UserID :" & DBList1.Text & "<br>Date :" & Date & "<br>Time :" & Time & "<br><hr WIDTH=""100%""></center><br>&nbsp;"

With Data1
        .DatabaseName = App.Path & "\login.exp"
        .RecordSource = "Test_Assign"
        .Refresh
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
     If .Recordset.Fields("UserID") = DBList1.Text Then
''''''''''''''''''''''''''''''''''''
table_body = table_body & "<tr><td>" & .Recordset.Fields("test_name") & "</td>"
     If .Recordset.Fields("attempt") Then
     no_test = no_test + 1
     total_p = total_p + .Recordset.Fields("score_p")
     
table_body = table_body & "<td>YES</td>" & "<td>" & .Recordset.Fields("score_p") & "</td>" & "<td>-</td>"
marquee = "<p align=""center""><marquee width=""50%"">Average Score %=" & (total_p / no_test) & "</marquee><br>&nbsp; </p>"
     Else
table_body = table_body & "<td>NO</td>" & "<td>-</td>" & "<td>" & .Recordset.Fields("e_date") & "</td>"
     End If
          
  '''''''''''''   MsgBox .Recordset.Fields("file_path")
    End If
    .Recordset.MoveNext
    Loop
        
End With


Open App.Path & "\" & DBList1.Text & ".htm" For Output As #2
Print #2, headder
Print #2, marquee
Print #2, table_body
Close 2
ShellExecute Me.hwnd, vbNullString, App.Path & "\" & DBList1.Text & ".htm", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
With datlogin
        .DatabaseName = App.Path & "\login.exp"
        .RecordSource = "SELECT UserID, Password, FirstName, LastName " & _
                                "FROM Login " & _
                              "WHERE Instructor = False"
        .Refresh
        End With
DBList1.ListField = "UserID"
End Sub
