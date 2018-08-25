VERSION 5.00
Begin VB.Form frmtest_paper 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   Picture         =   "frmtest_paper.frx":0000
   ScaleHeight     =   7620
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdjump 
      Caption         =   "Jump Question"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmtest_paper.frx":450D
      MousePointer    =   99  'Custom
      Picture         =   "frmtest_paper.frx":4817
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer q_time 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   7080
   End
   Begin VB.Data datbank 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer second 
      Interval        =   1000
      Left            =   720
      Top             =   7080
   End
   Begin VB.Data datfind 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dattest 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmd_abort 
      Caption         =   "Abort Test"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MouseIcon       =   "frmtest_paper.frx":4E14
      MousePointer    =   99  'Custom
      Picture         =   "frmtest_paper.frx":511E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "    Warning : if you ABORT the test now all Question left will be considered as NO ATTEMPT and results will be stored      "
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmd_next 
      Caption         =   "Next Question"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      MouseIcon       =   "frmtest_paper.frx":571B
      MousePointer    =   99  'Custom
      Picture         =   "frmtest_paper.frx":5A25
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1935
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H80000016&
      Caption         =   " "
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   5880
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000016&
      Caption         =   " "
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   5400
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000016&
      Caption         =   " "
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   4920
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000016&
      Caption         =   " "
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000016&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmtest_paper.frx":6022
      Top             =   5400
      Width           =   5415
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmtest_paper.frx":6024
      Top             =   4920
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmtest_paper.frx":6026
      Top             =   4440
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmtest_paper.frx":6028
      Top             =   3960
      Width           =   5415
   End
   Begin VB.TextBox txt_q 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmtest_paper.frx":602A
      Top             =   2760
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2640
      Picture         =   "frmtest_paper.frx":602C
      Top             =   120
      Width           =   2700
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer B"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   34
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Answer D"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   33
      Top             =   5400
      Width           =   780
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Answer C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   32
      Top             =   4920
      Width           =   765
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q="
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   31
      Top             =   2760
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   30
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lb_qno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   29
      Top             =   2760
      Width           =   105
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7680
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label lb_t 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5880
      TabIndex        =   28
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO ATTEMPT  (Choose this option to avoid Negative Marking)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2520
      TabIndex        =   27
      Top             =   5880
      Width           =   4800
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question Time :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   26
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lb_test 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5880
      TabIndex        =   25
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label lb_q 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5880
      TabIndex        =   24
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label lb_m 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5880
      TabIndex        =   23
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label lb_user 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1920
      TabIndex        =   22
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label lb_qt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2520
      TabIndex        =   21
      Top             =   2040
      Width           =   105
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Marks :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   20
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4320
      TabIndex        =   19
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label lb_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0/0"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   18
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Marks :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4320
      TabIndex        =   17
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   16
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Questions :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4320
      TabIndex        =   15
      Top             =   1800
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   14
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Name :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4320
      TabIndex        =   13
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lb_total 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2520
      TabIndex        =   8
      Top             =   2280
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      Height          =   5655
      Left            =   120
      Top             =   1320
      Width           =   7575
   End
End
Attribute VB_Name = "frmtest_paper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim long_sec As Boolean
Dim total_m As Integer
Dim sec As Integer
Dim total As Integer
Dim total_p As Integer
Dim total_q As Integer
Dim user_ans() As String
Dim table_body As String
Private Sub jj()
Do Until jump < 1

Loop
End Sub

Private Sub report()
Dim headder As String
Dim marquee As String
Dim assby As String
Dim pp As Integer
total_p = (total * 100) / total_m

headder = "<html><head><title>Test Result Report</title></head>" _
& "<center><b><u><font size=+1>TEST RESULT REPORT</font></u></b>" _
& "<br>Student Name :" & frmLogin.user_name & "<br>UserID :" & frmLogin.user_id & "<br>Test Name:" & frmLogin.test_name _
& "<br>Total Marks Scored :" & total & "<br>Percentage :" & total_p & "%<br><hr WIDTH=""100%""></center><br>&nbsp;"

With datbank
.DatabaseName = frmLogin.bank_filename
.RecordSource = "all_test"
.Refresh
.Recordset.FindFirst ("test_name='" & frmLogin.test_name & "'")
assby = .Recordset.Fields("User_ID")
pp = .Recordset.Fields("pp")
End With

If total_p >= pp Then
marquee = "<p align=""center""><marquee width=""50%"">Congratulations!!! You have PASSED this Test</marquee><br>&nbsp; </p>"
Else
marquee = "<p align=""center""><marquee width=""50%"">Sorry!!! You have FAILED in this Test</marquee><br>&nbsp; </p>"
End If
'''''''''''''''''''''''''''''''''''''''''''''
Open App.Path & "\" & frmLogin.user_id & frmLogin.test_name & ".htm" For Output As #1
Print #1, headder
Print #1, marquee
Print #1, table_body

Close 1
ShellExecute Me.hwnd, vbNullString, App.Path & "\" & frmLogin.user_id & ".htm", vbNullString, "C:\", SW_SHOWNORMAL
End Sub


Private Sub load_q()

Option5.Value = True
sec = dattest.Recordset.Fields("time")
lb_qt.Caption = sec
If sec = 0 + 1 Then
cmdjump.Visible = False
second.Enabled = False
q_time.Enabled = False
lb_qt.Caption = "Unlimited"
Else
    If sec <= 30 Then
    long_sec = False
    q_time.Interval = 1000 * sec
    Else
    sec = sec - 30
    q_time.Interval = 1000 * 30
    long_sec = True
    End If
'q_time.Interval = 1000 * sec
'cmdjump.Visible = True
q_time.Enabled = True
second.Enabled = True

lb_qt.Visible = True
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 lb_qno.Caption = dattest.Recordset.AbsolutePosition + 1


datbank.Recordset.FindFirst ("q_id=" & dattest.Recordset.Fields("q_id"))
txt_q = datbank.Recordset.Fields("question")
Text1 = datbank.Recordset.Fields("optA")
Text2 = datbank.Recordset.Fields("optB")
Text3 = datbank.Recordset.Fields("optC")
Text4 = datbank.Recordset.Fields("optD")
If datbank.Recordset.Fields("type") = "T" Then
Option3.Enabled = False
Option4.Enabled = False
Else
Option3.Enabled = True
Option4.Enabled = True
End If

End Sub
Private Sub ans()

Dim real_ans As String
real_ans = datbank.Recordset.Fields("answer")
''''''''''''''''''''''''
table_body = table_body & "<tr><td>" & lb_qno.Caption & "</td>" & "<td>" & real_ans & "</td>"
''''''''''''''''''''''''
user_ans(dattest.Recordset.AbsolutePosition) = "E"
If Option5.Value = True Then
table_body = table_body & "<td>" & "NOT ATTEMPTED" & "</td>" & "<td>" & "0" & "</td>"
Exit Sub
Else

If Option1.Value Then table_body = table_body & "<td>" & "A" & "</td>"
If Option2.Value Then table_body = table_body & "<td>" & "B" & "</td>"
If Option3.Value Then table_body = table_body & "<td>" & "C" & "</td>"
If Option4.Value Then table_body = table_body & "<td>" & "D" & "</td>"

If (real_ans = "A" And Option1.Value = True) Or (real_ans = "B" And Option2.Value = True) Or (real_ans = "C" And Option3.Value = True) Or (real_ans = "D" And Option4.Value = True) Then
table_body = table_body & "<td>" & "+" & dattest.Recordset.Fields("p_marks") & "</td>"
total = total + dattest.Recordset.Fields("p_marks")
Else
table_body = table_body & "<td>" & "-" & dattest.Recordset.Fields("n_marks") & "</td>"
total = total - dattest.Recordset.Fields("n_marks")
End If
lb_total.Caption = total
End If

End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub cmd_abort_Click()
total_p = (total * 100) / total_m
datbank.Recordset.Close
MsgBox "You have Aborted the Test:" & frmLogin.test_name & ",Express Test will now create a .HTM result Report and display it invoking you default browser", , "Test Aborted!!!"
report

With datbank
        .DatabaseName = App.Path & "\login.exp"
        .RecordSource = "select * from Test_Assign where UserID='" & frmLogin.user_id & "' and test_name ='" & frmLogin.test_name & "'"
        .Refresh
        .Recordset.Edit
        .Recordset.Fields("attempt") = True
        .Recordset.Fields("score_p") = total_p
        .Recordset.Fields("a_date") = Date
        .UpdateRecord
    End With
    ''''''''''''''''''''''''''''''
    'back to login form'''''''''''
    ''''''''''''''''''''''''''''''
    Load frmLogin
    frmLogin.Show
    Unload Me
End Sub

Private Sub cmd_next_Click()
q_time_Timer
second_Timer

End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''
'draw form
Dim rgn As Long, rgn2 As Long
Dim tmp As Long
Dim x As Integer, y As Integer
x = 176
y = 10
rgn = CreateRoundRectRgn(x, y, x + 181, y + 75, 30, 30)
x = 0
y = 10 + 70
rgn2 = CreateRoundRectRgn(x, y, x + 523, y + 400, 25, 25)
tmp = CombineRgn(rgn, rgn, rgn2, 2)
'set the window
tmp = SetWindowRgn(Me.hwnd, rgn, True)


'''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim total_t As Integer
table_body = "<center><table BORDER=5 COLS=4 WIDTH=""100%"" ><tr><td>Question no</td><td>Correct Answer</td><td>Student's Answer</td><td>Marks Awarded</td></tr>"

Show
lb_date.Caption = Date
lb_user.Caption = frmLogin.user_id
lb_test.Caption = frmLogin.test_name

datbank.DatabaseName = frmLogin.bank_filename
datbank.RecordSource = "q_bank"
datbank.Refresh

With dattest
datfind.DatabaseName = frmLogin.bank_filename
        .DatabaseName = frmLogin.bank_filename
        .RecordSource = frmLogin.test_name
        .Refresh
        If dattest.Recordset.AbsolutePosition < 0 Then
        MsgBox "This Test is Empty"
        End
        Else
.Recordset.MoveFirst
.Recordset.MoveLast
total_q = .Recordset.AbsolutePosition
lb_q.Caption = total_q + 1
ReDim user_ans(total_q + 1) As String
.Recordset.MoveFirst
    Do Until .Recordset.EOF
    datfind.RecordSource = "select * from q_bank where q_bank.q_id=" & .Recordset.Fields("q_id")
    datfind.Refresh
    total_m = total_m + .Recordset.Fields("p_marks")
    total_t = total_t + .Recordset.Fields("time")
    .Recordset.MoveNext
    Loop
    lb_m.Caption = total_m
    lb_t.Caption = total_t
    .Recordset.MoveFirst

End If
End With
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'important NOTE:- each question time is set to q_timer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
load_q
'If sec <= 30 Then
'long_sec = False
'q_time.Interval = 1000 * sec
'Else
'sec = sec - 30
'q_time.Interval = 1000 * 30
'long_sec = True
'End If
'q_time.Enabled = True
End Sub

Private Sub Option1_Click()
user_ans(dattest.Recordset.AbsolutePosition) = "A"

End Sub

Private Sub Option2_Click()
user_ans(dattest.Recordset.AbsolutePosition) = "B"
End Sub

Private Sub Option3_Click()
user_ans(dattest.Recordset.AbsolutePosition) = "C"
End Sub

Private Sub Option4_Click()
user_ans(dattest.Recordset.AbsolutePosition) = "D"
End Sub

Private Sub Option5_Click()
user_ans(dattest.Recordset.AbsolutePosition) = "E"
End Sub

Private Sub q_time_Timer()

''''to solve the timer.interval over flow problem
If long_sec Then

    If sec <= 30 Then
    long_sec = False
    q_time.Interval = 1000 * sec
    Else
    sec = sec - 30
    q_time.Interval = 1000 * sec
    long_sec = True
    End If
Exit Sub

End If
'''''''''''''''''''''''''''''''''''''''''''''

If dattest.Recordset.AbsolutePosition >= total_q Then

MsgBox "You have completed the Test:" & frmLogin.test_name & ", Express Test will now create a .HTM result Report and display it invoking you default browser", , "Test Finished!!!"
datbank.Recordset.Close
report

With datbank
        .DatabaseName = App.Path & "\login.exp"
        .RecordSource = "select * from Test_Assign where UserID='" & frmLogin.user_id & "' and test_name ='" & frmLogin.test_name & "'"
        .Refresh
        .Recordset.Edit
        .Recordset.Fields("attempt") = True
        .Recordset.Fields("score_p") = total_p
        .Recordset.Fields("a_date") = Date
        .UpdateRecord
    End With
End
Else
ans
dattest.Recordset.MoveNext
load_q

End If
End Sub

Private Sub second_Timer()

'sec = sec - 1
lb_qt.Caption = Val(lb_qt.Caption) - 1
'MsgBox sec

End Sub
