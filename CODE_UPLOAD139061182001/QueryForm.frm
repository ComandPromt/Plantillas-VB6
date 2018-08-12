VERSION 5.00
Begin VB.Form QueryForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query Builder"
   ClientHeight    =   3810
   ClientLeft      =   4575
   ClientTop       =   3525
   ClientWidth     =   5415
   Icon            =   "QueryForm.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "QueryForm.frx":0442
   ScaleHeight     =   3810
   ScaleWidth      =   5415
   Begin VB.ComboBox AndOR 
      Height          =   315
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2452
      Width           =   630
   End
   Begin VB.CommandButton Wildcards 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Wildcards"
      Height          =   255
      Left            =   2235
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   1
      Left            =   2955
      TabIndex        =   9
      Top             =   2804
      Width           =   2235
   End
   Begin VB.ComboBox Field 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2804
      Width           =   2190
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   0
      Left            =   2955
      TabIndex        =   6
      Top             =   2160
      Width           =   2235
   End
   Begin VB.ComboBox Field 
      Height          =   315
      Index           =   0
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2100
      Width           =   2190
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   3540
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1350
      Left            =   2640
      TabIndex        =   4
      Top             =   255
      Width           =   2655
      Begin VB.CheckBox chkOpen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show only open jobs"
         Height          =   225
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sort By"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   270
      Width           =   2175
      Begin VB.OptionButton OptCompanyName 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Company Name"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optContact 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contact"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton OK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
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
      Left            =   2610
      TabIndex        =   15
      Top             =   2804
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
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
      Left            =   2610
      TabIndex        =   14
      Top             =   2137
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show only records WHERE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1072
      TabIndex        =   13
      Top             =   1710
      Width           =   3270
   End
End
Attribute VB_Name = "QueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As Connection
Dim rs As ADODB.Recordset
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cancel.FontBold = True
    OK.FontBold = False
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Field_Change(Index As Integer)
    If Index = 0 And Field(0).ListIndex <> -1 Then Field(1).Enabled = True
End Sub

Private Sub Field_Click(Index As Integer)
    If Index = 0 And Field(0).ListIndex <> -1 Then Field(1).Enabled = True
End Sub

Private Sub Form_Load()

    Dim i As Integer
    Dim sql As String
    Dim connstring As String

    connstring = "DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=" & App.Path & "\recruiter.mdb;"
    Set conn = New ADODB.Connection
    conn.Open (connstring)
    sql = "select * from QueryStrings"
    Set rs = New ADODB.Recordset
    
    'use the recordset open method instead of the connection.execute method
    'execute creates forward only recordsets. We want to navigate back and forth
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic '1 = adopenKeyset, 3 = adlockoptimistic

    
    
    Me.Width = 5550
    Me.Height = 4200
    CenterForm Me
    OptCompanyName.Value = True
    
    'load combos on query field
    For i = 0 To 1
        With QueryForm.Field(i)
            .AddItem "Company Name"
            .AddItem "Contact"
            .AddItem "Phone"
            .AddItem "Email"
            .AddItem "Website"
            .AddItem "Address"
            .AddItem "City"
            .AddItem "State"
            .AddItem "Zip"
            .AddItem "Rating"
            .AddItem "Comments"
            .AddItem "Still Open"
            .AddItem "Job Description"
            .AddItem "First Contact"
            .AddItem "Follow up Date"
            .AddItem "Fax"
            .AddItem "Referred By"
            .ListIndex = -1
        End With
    Next
    

    With AndOR
        .AddItem "OR"
        .AddItem "AND"
        .ListIndex = 0
    End With

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OK.FontBold = False
    Cancel.FontBold = False
    Wildcards.FontBold = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
End Sub

Private Sub OK_Click()
    Dim sqlstring As String
    
    'variables to hold the field values used in the query
    Dim Fields(1) As String
    Dim Values(1) As String
    Dim AndOrs As String
    'troubleshooting variable
    Dim where As Integer
    
    'variable for the data environment
    Dim DE1 As New DE
    Dim i As Integer
    
    On Error GoTo QError

    'check to see if we want just open jobs and start query string
    If chkOpen.Value = vbChecked Then
        sqlstring = "SELECT * FROM Info WHERE StillOpen = Yes"
    Else
        sqlstring = "SELECT * FROM Info"
    End If
   
   'fill the variables with the field values
    AndOrs = AndOR.Text
  
    For i = 0 To 1
        Select Case Field(i).ListIndex
            Case 0
                Fields(i) = "CompanyName"
            Case 1
                Fields(i) = "Contact"
            Case 2
                Fields(i) = "Phone"
            Case 3
                Fields(i) = "Email"
            Case 4
                Fields(i) = "Website"
            Case 5
                Fields(i) = "Address"
            Case 6
                Fields(i) = "City"
            Case 7
                Fields(i) = "State"
            Case 8
                Fields(i) = "Zip"
            Case 9
                Fields(i) = "Rating"
            Case 10
                Fields(i) = "Comments"
            Case 11
                Fields(i) = "StillOpen"
            Case 12
                Fields(i) = "JobDescription"
            Case 13
                Fields(i) = "FirstDate"
            Case 14
                Fields(i) = "NextDate"
            Case 15
                Fields(i) = "Fax"
            Case 16
                Fields(i) = "ReferredBy"
            Case Else
                Fields(i) = ""
        End Select
        Values(i) = Value(i).Text
    Next
    
    'if there are any field values then add them to the query string
    If Trim(Fields(0)) <> "" Then
        If chkOpen.Value = vbChecked Then
            sqlstring = sqlstring & " AND "
        Else
            sqlstring = sqlstring & " WHERE "
        End If
        'don't use the like keyword with numeric values
        If Trim(Fields(0)) = "StillOpen" Or Trim(Fields(0)) = "Rating" Then
            sqlstring = sqlstring & Trim(Fields(0)) & " = " & Trim(Values(0))
        ElseIf Trim(Fields(0)) = "FirstDate" Or Trim(Fields(0)) = "NextDate" Then
            sqlstring = sqlstring & Trim(Fields(0)) & "= #" & Trim(Values(0)) & "#"
        Else
            sqlstring = sqlstring & Trim(Fields(0)) & " LIKE '" & Trim(Values(0)) & "'"
        End If
    End If
    If Trim(Fields(1)) <> "" Then
        'don't use the like keyword with numeric values
        If Trim(Fields(1)) = "StillOpen" Or Trim(Fields(1)) = "Rating" Then
            sqlstring = sqlstring & " " & AndOrs & " " & Trim(Fields(1)) & " = " & Trim(Values(1))
        ElseIf Trim(Fields(1)) = "FirstDate" Or Trim(Fields(1)) = "NextDate" Then
            sqlstring = sqlstring & " " & AndOrs & " " & Trim(Fields(1)) & "= #" & Trim(Values(1)) & "#"
        Else
            sqlstring = sqlstring & " " & AndOrs & " " & Trim(Fields(1)) & " LIKE '" & Trim(Values(1)) & "'"
        End If
    End If
    
    'add sort to the query string if needed
    If OptCompanyName = True Then
        sqlstring = sqlstring & " ORDER BY CompanyName"
    Else
        sqlstring = sqlstring & " ORDER BY Contact"
    End If
    
    'change recordsource to the new query string and refresh
    RecruiterForm.info.RecordSource = sqlstring
    RecruiterForm.info.Refresh
    
    'reset error here and make a new error handler for write to the
    'querystring table
    On Error GoTo 0
    
'************************************************************************
    'open the command object for the query strings table and
    'check to see if this query is already in it, if not
    'then add a new record.
    'Before saving the string we need to replace single quotes contained
    'in the string with another character. We'll use chr(0). The single
    'quotes cause trouble in the find function. We will convert them back before we
    'show them in the saved query form text box
'************************************************************************
    On Error GoTo WriteError
    where = 0
    i = InStr(1, sqlstring, "'")
    If i <> 0 Then sqlstring = Replace(sqlstring, "'", Chr(0))
    where = 1
    
    If rs.EOF And rs.BOF Then
    where = 2
    Else
        rs.MoveFirst
        where = 3
        Do While Not rs.EOF
            where = 4
            If UCase(rs.Fields("SQLSTRING")) = UCase(sqlstring) Then
                where = 5
                Exit Do
            End If
            where = 6
            rs.MoveNext
        Loop
        'rs.Find "SqlString ='" & sqlstring & "'"
    End If
    where = 7
    If rs.EOF = True Then
        where = 8
        rs.AddNew
        where = 9
        rs.Fields("sqlstring") = sqlstring
        where = 10
        rs.Update
        where = 11
    End If
    Unload Me
    
    Exit Sub
QError:
     MsgBox "There was an error processing your query. Please make sure  " & vbNewLine & _
            "that you used valid wildcard characters, and try again." & vbNewLine & _
            "Error Number " & Err.Number & vbNewLine & Err.Description, vbCritical, "Query Error"
    Unload Me
    
WriteError:
    MsgBox "Problem saving Query String. Error Number " & Err.Number & vbNewLine & _
        Err.Description & " From " & Err.Source & " Where = " & where
    Unload Me
End Sub

Private Sub OK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OK.FontBold = True
    Cancel.FontBold = False
End Sub

Private Sub Wildcards_Click()
        MsgBox "You can use the following wildcard characters on all fields but 'Still Open' " _
        & vbNewLine & "'_'  replaces any single character." & vbNewLine & _
        "'%'  replaces any string of zero or more characters." & vbNewLine & _
        "'[ ]'  will choose a range of characters ex.[5-8] will match 5, 6, 7,or 8. " & _
        vbNewLine & "'[^]' will exclude any character you place after the caret(^) character." _
        , , "Acceptable Wildcard Characters"
End Sub


Private Sub Wildcards_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Wildcards.FontBold = True
End Sub
