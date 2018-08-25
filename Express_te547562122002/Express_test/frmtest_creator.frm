VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmtest_creator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Express Test Creator"
   ClientHeight    =   7260
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9930
   Icon            =   "frmtest_creator.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_ppset 
      Caption         =   "SET"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   6000
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1815
      LargeChange     =   5
      Left            =   4680
      Max             =   5
      Min             =   100
      SmallChange     =   5
      TabIndex        =   25
      Top             =   5160
      Value           =   35
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox list_id 
      Height          =   3180
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox list_q 
      Height          =   3180
      ItemData        =   "frmtest_creator.frx":0442
      Left            =   960
      List            =   "frmtest_creator.frx":0444
      TabIndex        =   19
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Data datfind 
      Caption         =   "Data find"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data dattest 
      Caption         =   "test"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton cmdedit_qb 
      Caption         =   "Edit Question Bank"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmd_test_del 
      Caption         =   "Delete Question"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mark Options"
      Height          =   1935
      Left            =   2040
      TabIndex        =   3
      Top             =   5040
      Width           =   2415
      Begin VB.CheckBox Check1 
         Caption         =   "Enable Negative Marking"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt_neg 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt_pos 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Negative"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Positive"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time Options"
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
      Begin VB.TextBox txt_time 
         Height          =   285
         Left            =   600
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "20"
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton opt_limit 
         Caption         =   "Limited Time"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt_unlmit 
         Caption         =   "Unlimited Time"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Sec"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdmove_q 
      Caption         =   ">>Move Question to Test Paper >>>>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   4440
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   6480
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Data datbank 
      Caption         =   "Data bank"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label9 
      Caption         =   "Set Pass Percentage"
      Height          =   435
      Left            =   5160
      TabIndex        =   29
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   195
      Left            =   5520
      TabIndex        =   27
      Top             =   5760
      Width           =   120
   End
   Begin VB.Label lb_pp 
      AutoSize        =   -1  'True
      Caption         =   "35"
      Height          =   195
      Left            =   5160
      TabIndex        =   26
      Top             =   5760
      Width           =   90
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   6960
      Picture         =   "frmtest_creator.frx":0446
      Top             =   0
      Width           =   2700
   End
   Begin VB.Label Label_t 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   7920
      TabIndex        =   24
      Top             =   6120
      Width           =   90
   End
   Begin VB.Label Label_m 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   7920
      TabIndex        =   23
      Top             =   6360
      Width           =   90
   End
   Begin VB.Label Label_q 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   7920
      TabIndex        =   22
      Top             =   5880
      Width           =   90
   End
   Begin VB.Label Label_name 
      AutoSize        =   -1  'True
      Caption         =   "No Name"
      Height          =   195
      Left            =   7920
      TabIndex        =   21
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total Time :"
      Height          =   195
      Left            =   7080
      TabIndex        =   18
      Top             =   6120
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total Marks :"
      Height          =   195
      Left            =   6960
      TabIndex        =   17
      Top             =   6360
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Question :"
      Height          =   195
      Left            =   6795
      TabIndex        =   16
      Top             =   5880
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Test Paper Name :"
      Height          =   195
      Left            =   6585
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6480
      Picture         =   "frmtest_creator.frx":198B
      Top             =   5160
      Width           =   480
   End
   Begin VB.Menu men_file 
      Caption         =   "&File"
      Begin VB.Menu men_qb 
         Caption         =   "New Question Bank"
      End
      Begin VB.Menu men_open_qb 
         Caption         =   "Open Question Bank"
      End
      Begin VB.Menu u 
         Caption         =   "-"
      End
      Begin VB.Menu men_tp 
         Caption         =   "New Test Paper"
      End
      Begin VB.Menu men_open_tp 
         Caption         =   "Open Test Paper"
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
      Begin VB.Menu men_ass 
         Caption         =   "Assign Test"
      End
      Begin VB.Menu p 
         Caption         =   "-"
      End
      Begin VB.Menu men_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu men_login 
      Caption         =   "&Login"
      Begin VB.Menu men_st_log 
         Caption         =   "New Student Login"
      End
      Begin VB.Menu men_st_log_edit 
         Caption         =   "Edit Student Login"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu men_te_log 
         Caption         =   "New Instructor Login"
      End
      Begin VB.Menu men_edit_your 
         Caption         =   "Edit Your Login"
      End
   End
   Begin VB.Menu men_rep 
      Caption         =   "&Reports"
      Begin VB.Menu men_sp 
         Caption         =   "Student Report"
      End
      Begin VB.Menu men_ip 
         Caption         =   "Instructor Report"
      End
      Begin VB.Menu pp 
         Caption         =   "-"
      End
      Begin VB.Menu men_t_report 
         Caption         =   "Test Report"
      End
   End
   Begin VB.Menu men_about 
      Caption         =   "&About"
      Begin VB.Menu men_about_eqb 
         Caption         =   "About Express Test"
      End
   End
End
Attribute VB_Name = "frmtest_creator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bank_filename As String
Public test_name As String


Public Sub load_list()
list_q.Clear
list_id.Clear

 If datbank.Recordset.AbsolutePosition < 0 Then
        MsgBox "This Question Bank is Empty"
        List1.Clear
        Exit Sub
        End If
    ''''''''''''''''''''''''''''''''''''
    datbank.Recordset.MoveFirst
    Do Until datbank.Recordset.EOF
    list_q.AddItem (datbank.Recordset.Fields("question"))
    datbank.Recordset.MoveNext
    Loop
    '''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''
    datbank.Recordset.MoveFirst
    Do Until datbank.Recordset.EOF
    list_id.AddItem (datbank.Recordset.Fields("q_id"))
    datbank.Recordset.MoveNext
    Loop
    '''''''''''''''''''''''''''''''''''''''
End Sub
Public Function load_bank(f_name As String)
bank_filename = f_name
With datbank
        .DatabaseName = bank_filename
        .RecordSource = "q_bank"
        .Refresh
    End With
End Function
Public Function load_test(t_name As String)
Dim total_m As Integer
Dim total_t As Integer
List1.Clear
test_name = t_name
With dattest
        datfind.DatabaseName = bank_filename 'new
        .DatabaseName = bank_filename
        .RecordSource = test_name
        .Refresh
        If .Recordset.AbsolutePosition < 0 Then
        MsgBox "This Test is Empty"
        GoTo jump
        End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   .Recordset.MoveFirst
   .Recordset.MoveLast
   
   Label_q.Caption = .Recordset.AbsolutePosition + 1
    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'MsgBox .Recordset.AbsolutePosition
    List1.Clear
    If .Recordset.AbsolutePosition >= 0 Then
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
    datfind.RecordSource = "select question from q_bank where q_bank.q_id=" & .Recordset.Fields("q_id")
    datfind.Refresh
    List1.AddItem (datfind.Recordset.Fields("question"))
    total_m = total_m + .Recordset.Fields("p_marks")
    total_t = total_t + .Recordset.Fields("time")
    .Recordset.MoveNext
    Loop
    Label_m.Caption = total_m
    Label_t.Caption = total_t
    End If
End With
jump:
Label_name.Caption = test_name
cmdmove_q.Enabled = True
cmd_test_del.Enabled = True
men_ass.Enabled = True
cmd_ppset.Enabled = True
End Function

Private Sub Check1_Click()
If txt_neg.Enabled Then
txt_neg = "0"
txt_neg.Enabled = False
Else
txt_neg = "0.5"
txt_neg.Enabled = True
End If

End Sub

Private Sub cmd_ppset_Click()
With datfind
    .DatabaseName = bank_filename
    .RecordSource = "all_test"
    .Refresh
    .Recordset.FindFirst ("test_name='" & test_name & "'")
    .Recordset.Edit
    .Recordset.Fields("pp") = VScroll1.Value
    .UpdateRecord
    End With
End Sub

Private Sub cmd_test_del_Click()
If List1.Text = "" Then
Exit Sub
Else
'MsgBox List1.ListIndex
With dattest
.Recordset.MoveFirst
Do Until .Recordset.AbsolutePosition = List1.ListIndex
.Recordset.MoveNext
Loop
.Recordset.Delete
List1.Clear
load_test (test_name)
End With
End If
End Sub

Private Sub cmdedit_qb_Click()
frmq_editor.Show (1)

End Sub

Private Sub cmdmove_q_Click()
If list_id.Text = "" Then
MsgBox "Please select a Question", , "Nothing to Move!"

Else
With datfind
.DatabaseName = bank_filename
.RecordSource = "Select * from q_bank where q_id = " & list_id.Text
.Refresh
End With

With dattest.Recordset
.AddNew
.Fields("q_id") = datfind.Recordset.Fields("q_id")
.Fields("time") = Val(txt_time)
.Fields("p_marks") = Val(txt_pos)
.Fields("n_marks") = Val(txt_neg)
End With
On Error GoTo err_dup

dattest.UpdateRecord
List1.AddItem (list_q.Text)
err_dup:
If Err = 3022 Then
MsgBox "Question already exists in the Test Paper", vbOKOnly, "Duplacate Question"
dattest.Recordset.CancelUpdate
'List1.Clear
Exit Sub
End If
Resume Next


End If

End Sub


Private Sub Form_Load()
'MsgBox "You have completed the Test:" & frmLogin.test_name & ",Express Test will now create a .HTM result Report and display it invoking you default browser", , "Test Finished!!!"
men_ass.Enabled = False
men_tp.Enabled = False
men_open_tp.Enabled = False
cmdedit_qb.Enabled = False
Show

End Sub

Private Sub list_id_Click()
list_q.ListIndex = list_id.ListIndex
End Sub

Private Sub list_q_Click()
'MsgBox list_q.ListCount
list_id.ListIndex = list_q.ListIndex

End Sub



Private Sub list_q_DblClick()
MsgBox list_q.Text, vbOKOnly, "Question Selected"

End Sub

Private Sub men_about_eqb_Click()
frmAbout.Show (1)


End Sub

Private Sub men_ass_Click()
frmAssign.Show (1)
End Sub

Private Sub men_edit_your_Click()
frmTeachPass.Show (1)

End Sub

Private Sub men_exit_Click()
'End
Load frmLogin
frmLogin.Show
Unload Me
End Sub

Private Sub men_ip_Click()
frmreport_i.Show (1)
End Sub

Private Sub men_open_qb_Click()
 With CDialog1
        .CancelError = False
        .InitDir = App.Path
        .Flags = 2
        .Filter = "Microsoft Access Files| *.mdb"
        .DialogTitle = "Open Question Bank"
        .ShowOpen
        
        If Dir(.FileName) <> "" Then
         bank_filename = .FileName
            'MsgBox "file found"
            Else
            MsgBox "Question Bank File NOT Found", vbCritical = vbOKOnly, "File Not Found"
            Exit Sub
        End If
If .FileName = "" Then Exit Sub

''''''''''''''''''''''''''''''''''''''''''''''
 'if the file is found condition has to be put
''''''''''''''''''''''''''''''''''''''''''''''

End With
load_bank (bank_filename)
load_list
men_tp.Enabled = True
men_open_tp.Enabled = True
cmdedit_qb.Enabled = True
End Sub

Private Sub men_open_tp_Click()
'load_test ("testing")
frmopen_test.Show (1)
End Sub

Private Sub men_qb_Click()
With CDialog1
        .CancelError = False
        .InitDir = App.Path
        .Flags = 2
        .Filter = "Microsoft Access Files| *.mdb"
        .DialogTitle = "New Question Bank"
        .ShowSave


If Dir(.FileName) <> "" And .FileName <> "" Then
''MsgBox "Question Bank File already exists ,Please delete it to create a File of same name "
Kill .FileName
ElseIf .FileName = "" Then
Exit Sub
End If
MsgBox .FileName
FileCopy App.Path & "\mdb.exp", .FileName
bank_filename = .FileName
End With
cmdedit_qb_Click
men_tp.Enabled = True
men_open_tp.Enabled = True
cmdedit_qb.Enabled = True
End Sub

Private Sub men_sp_Click()
frmreport_s.Show (1)
End Sub

Private Sub men_st_log_Click()
'load form and set caption to student
'Load frmAccount
    frmAccount.Caption = "New Student Login"
    frmAccount.Show (1)
    
End Sub

Private Sub men_st_log_edit_Click()

frmStudPass.Show (1)
End Sub

Private Sub men_t_report_Click()
frmreport_t.Show (1)
End Sub

Private Sub men_te_log_Click()
frmAccount.Caption = "New Instructor Login"
frmAccount.Show (1)
End Sub

Private Sub men_tp_Click()
test_name = InputBox("Enter Test Paper Name (max 50 char)", "New Test Paper")
If test_name = "" Then
Exit Sub
Else
    Dim NewBank As Database, MyWS As Workspace
    Dim T1 As TableDef
    Dim T1Flds(1 To 4) As Field
    Dim T1Idx As Index
    Dim myRec As Recordset
    Dim checkDIR As String
    
    
    'create new question bank database
        Set MyWS = DBEngine.Workspaces(0)
        Set NewBank = MyWS.OpenDatabase(bank_filename)
        
        Set T1 = NewBank.CreateTableDef(test_name)
        
        Set T1Flds(1) = T1.CreateField("q_id", dbInteger)
        Set T1Flds(2) = T1.CreateField("time", dbInteger)
        Set T1Flds(3) = T1.CreateField("p_marks", dbSingle)
        Set T1Flds(4) = T1.CreateField("n_marks", dbSingle)
        
        T1.Fields.Append T1Flds(1)
        T1.Fields.Append T1Flds(2)
        T1.Fields.Append T1Flds(3)
        T1.Fields.Append T1Flds(4)
        
        Set T1Idx = T1.CreateIndex("q_id")
        T1Idx.Primary = True
        T1Idx.Unique = True
        T1Idx.Required = True
        Set T1Flds(1) = T1Idx.CreateField("q_id")
        T1Idx.Fields.Append T1Flds(1)
        T1.Indexes.Append T1Idx
        NewBank.TableDefs.Append T1
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'test table created no insetr it into all_test table
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Dim all_t As Data
    With datfind
    .DatabaseName = bank_filename
    .RecordSource = "all_test"
    .Refresh
    .Recordset.AddNew
    .Recordset.Fields("test_name") = test_name
    .Recordset.Fields("date") = Date
    .Recordset.Fields("User_ID") = "newdfdfd" 'current_uid
    .Recordset.Fields("pp") = VScroll1.Value
    .UpdateRecord
    End With
    load_test (test_name)
    
    
    End If
End Sub

Private Sub opt_limit_Click()
txt_time = "20"
txt_time.Enabled = True
End Sub

Private Sub opt_unlmit_Click()
txt_time = "0"
txt_time.Enabled = False
End Sub

Private Sub txt_neg_Change()

If IsNumeric(txt_neg) = False Then
Beep
txt_neg = ""
ElseIf Val(txt_pos.Text) < Val(txt_neg.Text) Then
Beep
txt_neg = ""
End If
End Sub

Private Sub txt_neg_LostFocus()
If txt_neg = "" Then
Beep
txt_neg.SetFocus
End If
End Sub

Private Sub txt_pos_Change()
If IsNumeric(txt_pos) = False Then
Beep
txt_pos = ""
End If
End Sub

Private Sub txt_pos_LostFocus()
If txt_pos = "" Then
Beep
txt_pos.SetFocus
End If
End Sub

Private Sub txt_time_Change()
If IsNumeric(txt_time) = False Then
Beep
txt_time = ""
End If
End Sub

Private Sub txt_time_LostFocus()
If txt_time = "" Then
Beep
txt_time.SetFocus
End If
End Sub

Private Sub VScroll1_Change()
lb_pp.Caption = VScroll1.Value
End Sub
