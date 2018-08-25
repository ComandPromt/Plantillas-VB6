VERSION 5.00
Begin VB.Form frmq_editor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question Bank Editor"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmInstructor1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_ok 
      Caption         =   " O K"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   6600
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Correct Answer"
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   4935
      Begin VB.OptionButton Option1 
         Caption         =   "A"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   " B"
         Height          =   195
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   " C"
         Height          =   195
         Left            =   2880
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   " D"
         Height          =   195
         Left            =   3960
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Question Type"
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Width           =   4935
      Begin VB.OptionButton opt_type_m 
         Caption         =   "Multiple Choice"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opt_type_t 
         Caption         =   "True/False"
         Height          =   195
         Left            =   3360
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Data Data1 
      Caption         =   "< Previous Question                         Next Question >"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Width           =   4980
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmd_update 
      Caption         =   "Update"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "Add New"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmd_del 
      Caption         =   "Delete"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txt_optd 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Text            =   " "
      Top             =   3480
      Width           =   4095
   End
   Begin VB.TextBox txt_optc 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Text            =   " "
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox txt_optb 
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Text            =   " "
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox txt_opta 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Text            =   " "
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox txt_q 
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmInstructor1.frx":0442
      Top             =   960
      Width           =   4815
   End
   Begin VB.TextBox txt_id 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   2760
      Picture         =   "frmInstructor1.frx":0444
      Top             =   0
      Width           =   2700
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Option D"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   3600
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Option C"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Option B"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Option A"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Full Question Text "
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Question ID"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmq_editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim b_filename As String
Public Sub showdata()
On Error GoTo no_rec:
With Data1.Recordset

txt_id = .Fields("q_id").Value & ""
txt_q = .Fields("question").Value & ""
txt_opta = .Fields("optA").Value & ""
txt_optb = .Fields("optB").Value & ""
txt_optc = .Fields("optC").Value & ""
txt_optd = .Fields("optD").Value & ""

If .Fields("type").Value = "M" Then
opt_type_m.Value = True
ElseIf .Fields("type").Value = "T" Then
opt_type_t.Value = True
End If
'Print .Fields("answer").Value
If .Fields("answer").Value = "A" Then
Option1.Value = True
ElseIf .Fields("answer").Value = "B" Then
Option2.Value = True
ElseIf .Fields("answer").Value = "C" Then
Option3.Value = True
ElseIf .Fields("answer").Value = "D" Then
Option4.Value = True
End If

End With
no_rec:
If Err = 3021 Then
MsgBox "No Record to display"
Exit Sub

End If

End Sub






Private Sub cmd_add_Click()
'Data1.Enabled = False
cmd_save.Enabled = True

cmd_ok.Enabled = False
cmd_update.Enabled = False
cmd_del.Enabled = False
Data1.Enabled = False
txt_id = ""
txt_q = ""
txt_opta = ""
txt_optb = ""
txt_optc = ""
txt_optd = ""
opt_type_m.Value = True
Option1.Value = True
Data1.Recordset.AddNew

End Sub

Private Sub cmd_del_Click()
'On Error GoTo no_rec:
response = MsgBox("Are you sure you want to Delete this Question ?", vbYesNo, "Deleting")
If response = vbNo Then
Exit Sub
Else
Data1.Recordset.Delete
Data1.Refresh
End If
End Sub

Private Sub cmd_ok_Click()
frmtest_creator.load_bank (frmtest_creator.bank_filename)
frmtest_creator.load_list
Unload Me
End Sub

Private Sub cmd_save_Click()
If txt_id = "" Or txt_q = "" Then
MsgBox "Please Enter Question ID and Question text", vbOKOnly, "Invalid Data"
Exit Sub
End If
If opt_type_m.Value = True Then
If txt_optc = "" Or txt_optd = "" Then
MsgBox "Please Enter all answer Options", vbOKOnly, "Invalid Data"
Exit Sub
End If
End If

With Data1.Recordset

.Fields("q_id").Value = Val(txt_id)
.Fields("question").Value = Trim(txt_q)
.Fields("optA").Value = Trim(txt_opta)
.Fields("optB").Value = Trim(txt_optb)
.Fields("optC").Value = Trim(txt_optc)
.Fields("optD").Value = Trim(txt_optd)

If opt_type_m.Value Then
 .Fields("type").Value = "M"
ElseIf opt_type_t.Value Then
 .Fields("type").Value = "T"
End If
'Print .Fields("answer").Value
If Option1.Value Then
 .Fields("answer").Value = "A"
ElseIf Option2.Value Then
 .Fields("answer").Value = "B"
ElseIf Option3.Value Then
 .Fields("answer").Value = "C"
ElseIf Option4.Value Then
 .Fields("answer").Value = "D"
End If

End With
On Error GoTo err_dup

Data1.UpdateRecord
err_dup:
If Err = 524 Then
MsgBox "Question ID already exists,Please enter another value ", vbOKOnly, "Invalid Data"

Data1.Recordset.CancelUpdate
End If

Resume Next
cmd_save.Enabled = False
cmd_ok.Enabled = True
cmd_update.Enabled = True
cmd_del.Enabled = True
cmd_add.Enabled = True
Data1.Enabled = True
End Sub

Private Sub cmd_update_Click()
cmd_save.Enabled = True
Data1.Enabled = False
cmd_ok.Enabled = False
cmd_add.Enabled = False
cmd_del.Enabled = False

Data1.Recordset.Edit

End Sub





Private Sub Data1_Reposition()
showdata
End Sub

Private Sub Form_Load()
With Data1
        .DatabaseName = frmtest_creator.bank_filename
        .RecordSource = "q_bank"
        .Refresh
    End With
    
End Sub

Private Sub opt_type_m_Click()
Option3.Enabled = True
Option4.Enabled = True

txt_opta.Enabled = True
txt_optb.Enabled = True
txt_optc.Enabled = True
txt_optd.Enabled = True

End Sub

Private Sub opt_type_t_Click()
Option3.Enabled = False
Option4.Enabled = False
txt_opta = "True"
txt_optb = "False"
txt_optc = ""
txt_optd = ""

txt_opta.Enabled = False
txt_optb.Enabled = False
txt_optc.Enabled = False
txt_optd.Enabled = False



End Sub

Private Sub txt_id_Change()
If IsNumeric(txt_id) = False Then
txt_id = ""
End If
End Sub
