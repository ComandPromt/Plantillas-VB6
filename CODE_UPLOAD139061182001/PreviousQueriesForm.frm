VERSION 5.00
Begin VB.Form PreviousQueryForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saved Queries"
   ClientHeight    =   4005
   ClientLeft      =   4545
   ClientTop       =   2880
   ClientWidth     =   6390
   Icon            =   "PreviousQueriesForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "PreviousQueriesForm.frx":0442
   ScaleHeight     =   4005
   ScaleWidth      =   6390
   Begin VB.CommandButton Previous 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<"
      Height          =   255
      Left            =   2408
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Move Previous"
      Top             =   3075
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">"
      Height          =   255
      Left            =   3233
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Move Next"
      Top             =   3075
      Width           =   735
   End
   Begin VB.CommandButton Delete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Delete"
      Height          =   255
      Left            =   2408
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3375
      Width           =   735
   End
   Begin VB.CommandButton Save 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3233
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3375
      Width           =   735
   End
   Begin VB.CommandButton MoveFirst 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<"
      Height          =   255
      Left            =   1583
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Move Previous"
      Top             =   3075
      Width           =   735
   End
   Begin VB.CommandButton MoveLast 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">>"
      Height          =   255
      Left            =   4073
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Move Previous"
      Top             =   3075
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C&lose"
      Height          =   255
      Left            =   4898
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3375
      Width           =   735
   End
   Begin VB.CommandButton Add 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add"
      Height          =   255
      Left            =   1583
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3375
      Width           =   735
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   570
      Width           =   6000
   End
   Begin VB.CommandButton Run 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Run"
      Default         =   -1  'True
      Height          =   255
      Left            =   758
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3375
      Width           =   735
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4073
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3375
      Width           =   735
   End
   Begin VB.Label RecordLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "of records"
      Height          =   180
      Left            =   1853
      TabIndex        =   6
      Top             =   3765
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saved Queries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2153
      TabIndex        =   4
      Top             =   165
      Width           =   2085
   End
End
Attribute VB_Name = "PreviousQueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As ADODB.Connection
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub Add_Click()

    If rs.BOF And rs.EOF Then
    Else
        If rs.EditMode = adEditAdd Then
            Call Save_Click
        End If
    End If
    rs.AddNew
    Add.Enabled = False
    Save.Enabled = True
    Cancel.Enabled = True
    
End Sub

Private Sub Add_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Add.FontBold = True
    Run.FontBold = False
    Delete.FontBold = False
End Sub

Private Sub Cancel_Click()

    rs.CancelUpdate
    Add.Enabled = True
    Save.Enabled = False
    Cancel.Enabled = False
    
End Sub

Private Sub Cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cancel.FontBold = True
    Save.FontBold = False
    cmdClose.FontBold = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdclose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdClose.FontBold = True
    Cancel.FontBold = False
End Sub



Private Sub Delete_Click()
    Dim response As Integer
    Dim bookmark As Variant
    
    'check to see if there are any records to delete
    If rs.EOF = True And rs.BOF = True Then
        txtSQL.Text = ""
        RecordLabel = "0 of 0 records"
        Exit Sub
    End If
    
    'check to see if the current record has been saved yet
    If rs.EditMode = adEditAdd Then
        response = MsgBox("The current record needs to be saved before you can perform this action. " & _
                vbNewLine & "Would you like to save the current record now?", vbYesNo + vbInformation, "Save Current Record?")
        If response = vbYes Then
            Call Save_Click
        Else
            Exit Sub
        End If
    End If
    
    rs.Delete
    
    'move off empty record
    'first check to see if there are any more records
    If rs.BOF = True And rs.EOF = True Then Exit Sub
    rs.MoveNext
    If rs.EOF = True And rs.BOF <> True Then
        rs.MovePrevious
    End If
    Exit Sub
    
    
DeleteError:
    MsgBox "Cannot delete " & Err.Number & " " & Err.Description & " " & Err.Source

End Sub

Private Sub Delete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Delete.FontBold = True
    Add.FontBold = False
    Save.FontBold = False
End Sub

Private Sub Form_Load()

    Dim sql As String
    Dim connstring As String
    '"DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=" & App.Path & "\recruiter.mdb;"
    connstring = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Recruiter.mdb"
    Set conn = New ADODB.Connection
    conn.Open (connstring)
    sql = "select * from QueryStrings"
    Set rs = New ADODB.Recordset
    
    'use the recordset open method instead of the connection.execute method
    'execute creates forward only recordsets. We want to navigate back and forth
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic    '1 = adopenKeyset, 3 = adlockoptimistic
    
'    Dim prop As Property
'    For Each prop In rs.Properties
'    If prop.Name = "Use Bookmarks" Or prop.Name = "Bookmarkable" Then
'        'rs.Close
'       prop.Value = True
'    End If
'    Debug.Print prop.Name & prop.Type & prop.Attributes & prop.Value
'    Next

    If rs.BOF <> True And rs.EOF <> True Then rs.MoveFirst
    CenterForm Me
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim ctrl As Control
    
    For Each ctrl In Controls
        If TypeOf ctrl Is CommandButton Then
            ctrl.FontBold = False
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
End Sub

Private Sub MoveFirst_Click()
    
    If rs.BOF = True And rs.EOF = True Then
        Beep
    Else
        'Check to see if an add is in process but not saved yet
        'if so save the record
        If rs.EditMode = adEditAdd Then
            Call Save_Click
        End If
        savedata
        rs.MoveFirst
    End If
End Sub

Private Sub MoveFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveFirst.FontBold = True
    Previous.FontBold = False
End Sub

Private Sub MoveLast_Click()

    If rs.BOF = True And rs.EOF = True Then
        Beep
    Else
        'Check to see if an add is in process but not saved yet
        'if so save the record
        If rs.EditMode = adEditAdd Then
            Call Save_Click
        End If
        savedata
        rs.MoveLast
    End If
End Sub

Private Sub MoveLast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveLast.FontBold = True
    cmdNext.FontBold = False
End Sub

Private Sub cmdNext_Click()
    
    If rs.EOF <> True Then
        'Check to see if an add is in process but not saved yet
        'if so save the record
        If rs.EditMode = adEditAdd Then
            Call Save_Click
        End If
        savedata
        rs.MoveNext
        If rs.EOF = True Then rs.MoveLast
    End If
End Sub


Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdNext.FontBold = True
    Previous.FontBold = False
    MoveLast.FontBold = False
End Sub

Private Sub Previous_Click()
    
    If rs.BOF <> True Then
        'Check to see if an add is in process but not saved yet
        'if so save the record
        If rs.EditMode = adEditAdd Then
            Call Save_Click
        End If
        savedata
        rs.MovePrevious
        If rs.BOF = True Then rs.MoveFirst
    End If
End Sub

Private Sub Previous_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Previous.FontBold = True
    MoveFirst.FontBold = False
    cmdNext.FontBold = False
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    On Error Resume Next
    
    'when the move is complete the record is updated
    Add.Enabled = True
    
    'check for any records and update record label
    If rs.BOF = True And rs.EOF = True Then
        txtSQL.Text = ""
        RecordLabel.Caption = "0 of 0 records"
        Exit Sub
    End If
    
    
    If rs.BOF = True And rs.EOF <> True Then
        Beep
        rs.MoveFirst
    ElseIf rs.EOF = True And rs.BOF <> True Then
        Beep
        rs.MoveLast
    ElseIf rs.BOF And rs.EOF Then
        txtSQL.Text = ""
        RecordLabel.Caption = "0 of 0 records"
        Exit Sub
    End If
    
    'bind txt box
    If IsNull(rs.Fields("SQLString")) Then
        txtSQL.Text = ""
    Else
        txtSQL.Text = Replace(rs.Fields("sqlstring"), Chr(0), "'")
    End If
    
    'update label
    RecordLabel.Caption = rs.AbsolutePosition & " of " & rs.RecordCount & " records"
    
End Sub

Private Sub Run_Click()
    
    Dim QueryString As String
    
    On Error Resume Next
    
    If rs.BOF And rs.EOF Then Exit Sub
    
    'check to see if the current record needs to be saved first
    If rs.EditMode = adEditAdd Then
        Call Save_Click
    End If

    If Trim(txtSQL.Text) <> "" Then
        QueryString = Trim(txtSQL.Text)
    End If
    
    'use querystring to redefine recordsource for recruiter form
    RecruiterForm.info.RecordSource = QueryString
    RecruiterForm.info.Refresh
    
    Unload Me
        
End Sub

Private Sub Run_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Run.FontBold = True
    Add.FontBold = False
End Sub

Private Sub Save_Click()
    Dim sqlstring As String
    
    sqlstring = Trim(txtSQL.Text)
    
    'replace single quotes with chr(0) to avoid problems with searches on this field
    sqlstring = Replace(sqlstring, "'", Chr(0))
    
    rs.Fields("sqlstring") = sqlstring
    rs.Update
    Add.Enabled = True
    Save.Enabled = False
    Cancel.Enabled = False
    
End Sub

Private Sub Save_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Save.FontBold = True
    Delete.FontBold = False
    Cancel.FontBold = False
End Sub

Private Sub savedata()
    Dim sqlstring As String
    
    sqlstring = Replace(txtSQL.Text, "'", Chr(0))

    rs.Fields("sqlstring") = sqlstring
End Sub
