VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RecruiterForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5265
   ClientLeft      =   2085
   ClientTop       =   2250
   ClientWidth     =   9570
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   5265
   ScaleWidth      =   9570
   Begin VB.TextBox ReferredBy 
      DataField       =   "ReferredBy"
      DataSource      =   "info"
      Height          =   285
      Left            =   5955
      TabIndex        =   14
      Top             =   2460
      Width           =   1815
   End
   Begin VB.TextBox txtWebsite 
      DataField       =   "Website"
      DataSource      =   "info"
      Height          =   285
      Left            =   5955
      TabIndex        =   13
      Top             =   2085
      Width           =   3390
   End
   Begin VB.CommandButton Go 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Website"
      Height          =   255
      Left            =   4875
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2100
      Width           =   1035
   End
   Begin VB.TextBox NextDate 
      DataField       =   "NextDate"
      DataSource      =   "info"
      Height          =   285
      Left            =   8310
      TabIndex        =   10
      Top             =   1350
      Width           =   1035
   End
   Begin VB.TextBox FirstDate 
      DataField       =   "FirstDate"
      DataSource      =   "info"
      Height          =   285
      Left            =   5955
      TabIndex        =   9
      Top             =   1350
      Width           =   1035
   End
   Begin VB.TextBox Fax 
      DataField       =   "Fax"
      DataSource      =   "info"
      Height          =   285
      Left            =   5955
      TabIndex        =   11
      Top             =   1725
      Width           =   3390
   End
   Begin VB.TextBox JobDescription 
      DataField       =   "JobDescription"
      DataSource      =   "info"
      Height          =   930
      Left            =   5955
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   270
      Width           =   3390
   End
   Begin VB.ComboBox Rating 
      DataField       =   "Rating"
      DataSource      =   "info"
      Height          =   315
      Left            =   8520
      TabIndex        =   15
      Top             =   2445
      Width           =   795
   End
   Begin VB.ComboBox cboState 
      DataField       =   "State"
      DataSource      =   "info"
      Height          =   315
      Left            =   5955
      TabIndex        =   16
      Top             =   2805
      Width           =   1575
   End
   Begin VB.TextBox txtZip 
      DataField       =   "Zip"
      DataSource      =   "info"
      Height          =   285
      Left            =   8520
      TabIndex        =   17
      Top             =   2820
      Width           =   795
   End
   Begin VB.TextBox txtCity 
      DataField       =   "City"
      DataSource      =   "info"
      Height          =   285
      Left            =   1290
      TabIndex        =   6
      Top             =   2820
      Width           =   3360
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataSource      =   "info"
      Height          =   285
      Left            =   1275
      TabIndex        =   5
      Top             =   2460
      Width           =   3390
   End
   Begin VB.CommandButton Wildcards 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Wildcards"
      Height          =   255
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3975
      Width           =   975
   End
   Begin VB.CommandButton Find 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Find"
      Height          =   255
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3435
      Width           =   975
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   5955
      TabIndex        =   19
      Top             =   3795
      Width           =   2250
   End
   Begin VB.ComboBox SearchField 
      Height          =   315
      Left            =   5955
      TabIndex        =   18
      Top             =   3435
      Width           =   2265
   End
   Begin VB.CommandButton FindNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Find Next"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3705
      Width           =   975
   End
   Begin VB.CheckBox StillOpen 
      Caption         =   "Still Open?"
      DataField       =   "StillOpen"
      DataSource      =   "info"
      Height          =   195
      Left            =   5985
      Picture         =   "Form1.frx":73D1
      TabIndex        =   24
      Top             =   4185
      Width           =   185
   End
   Begin VB.CommandButton Email 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Send Email"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2100
      Width           =   1080
   End
   Begin VB.CommandButton Close 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Close"
      Height          =   255
      Left            =   8535
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4425
      Width           =   855
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4425
      Width           =   855
   End
   Begin VB.CommandButton Save 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4425
      Width           =   855
   End
   Begin VB.CommandButton Delete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Delete"
      Height          =   255
      Left            =   6015
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4425
      Width           =   855
   End
   Begin VB.CommandButton Add 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add"
      Height          =   255
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4425
      Width           =   855
   End
   Begin VB.TextBox txtComments 
      DataField       =   "Comments"
      DataSource      =   "info"
      Height          =   1590
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3420
      Width           =   4485
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "Email"
      DataSource      =   "info"
      Height          =   285
      Left            =   1275
      TabIndex        =   4
      Top             =   2085
      Width           =   3390
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "Phone"
      DataSource      =   "info"
      Height          =   285
      Left            =   1275
      TabIndex        =   2
      Top             =   1725
      Width           =   3390
   End
   Begin VB.TextBox txtContact 
      DataField       =   "Contact"
      DataSource      =   "info"
      Height          =   285
      Left            =   1275
      TabIndex        =   1
      Top             =   1350
      Width           =   3390
   End
   Begin VB.TextBox txtCompanyName 
      DataField       =   "CompanyName"
      DataSource      =   "info"
      Height          =   285
      Left            =   1275
      TabIndex        =   0
      Top             =   285
      Width           =   3390
   End
   Begin MSAdodcLib.Adodc info 
      Height          =   330
      Left            =   5175
      Top             =   4665
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Recruiter.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Recruiter.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * FROM info Order by CompanyName"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referred by:"
      Height          =   195
      Left            =   5040
      TabIndex        =   45
      Top             =   2505
      Width           =   870
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      Height          =   195
      Left            =   5610
      TabIndex        =   44
      Top             =   1770
      Width           =   300
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Contact:"
      Height          =   195
      Left            =   4980
      TabIndex        =   43
      Top             =   1395
      Width           =   930
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Follow up Date:"
      Height          =   195
      Left            =   7140
      TabIndex        =   42
      Top             =   1395
      Width           =   1110
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Description:"
      Height          =   195
      Left            =   4770
      TabIndex        =   41
      Top             =   330
      Width           =   1140
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rating:"
      Height          =   195
      Left            =   7935
      TabIndex        =   40
      Top             =   2505
      Width           =   510
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      Height          =   195
      Left            =   5490
      TabIndex        =   39
      Top             =   2865
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
      Height          =   195
      Left            =   8175
      TabIndex        =   38
      Top             =   2865
      Width           =   270
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      Height          =   195
      Left            =   885
      TabIndex        =   37
      Top             =   2865
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   195
      Left            =   615
      TabIndex        =   36
      Top             =   2505
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Still Open?"
      Height          =   255
      Left            =   5160
      TabIndex        =   23
      Top             =   4155
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      Height          =   195
      Left            =   5610
      TabIndex        =   35
      Top             =   3840
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Field"
      Height          =   195
      Left            =   5025
      TabIndex        =   34
      Top             =   3465
      Width           =   885
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   33
      Top             =   330
      Width           =   1170
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   5
      Left            =   2220
      TabIndex        =   32
      Top             =   3180
      Width           =   780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   31
      Top             =   1770
      Width           =   510
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact:"
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   30
      Top             =   1395
      Width           =   600
   End
End
Attribute VB_Name = "RecruiterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************
'RecruiterApp is sort of an address book type app to keep track of
'**Jobs that you apply for that lets you add and delete contacts.
'**Keeps track of jobs that are StillPending (Highlights comments in red)
'**Go to website or send email to your contacts
'Using Access Database
'
'Author: Rick Bales copyright 2000
'rb.sb@gte.net
'Date: 12/21/2000
'**************************************************************************


Dim db As Connection
Dim bookmark As Variant
'for email and websites
Dim WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1
Dim WithEvents IE2 As InternetExplorer
Attribute IE2.VB_VarHelpID = -1
Dim Searching As Boolean
Dim DE1 As DE
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1


Private Sub Add_Click()
    
    On Error Resume Next
    'add a new record
    info.Recordset.AddNew
    
    'User has the choice of saving the new record or cancelling
    'all records can be 0 length. No need to test the fields
    Save.Enabled = True
    Cancel.Enabled = True
End Sub
Private Sub savedata()

    'assign all the variables to the database
    info.Recordset!CompanyName = Trim(txtCompanyName)
    info.Recordset!Contact = Trim(txtContact)
    info.Recordset!phone = Trim(txtPhone)
    info.Recordset!Email = Trim(txtEmail)
    info.Recordset!Comments = Trim(txtComments)
    info.Recordset!Website = Trim(txtWebsite)
    info.Recordset!StillOpen = StillOpen
End Sub



Private Sub Add_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Add.FontBold = True
    Delete.FontBold = False
End Sub

Private Sub Cancel_Click()
    
    On Error Resume Next
    info.Recordset.CancelUpdate
    
    'go back to record that you were on before the addnew was started.
    If info.Recordset.RecordCount > 1 Then info.Recordset.bookmark = bookmark
    Save.Enabled = False
    Cancel.Enabled = False
End Sub

Private Sub Cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cancel.FontBold = True
    Save.FontBold = False
    RecruiterForm.Close.FontBold = False
End Sub

Private Sub Delete_Click()
    Dim response As Integer
    'Use local variable so we don't get mixed up the
    Dim bookmark As Variant
    
    'first check to see if there are any more records
    If info.Recordset.BOF = True And info.Recordset.EOF = True Then Exit Sub
    
    On Error GoTo DeleteError
    response = MsgBox("Delete this record?", vbCritical + vbYesNo, "Delete?")
    If response = vbNo Then Exit Sub
    
    'save current record pointer
    If info.Recordset.RecordCount >= 1 Then
       bookmark = info.Recordset.bookmark
    End If
    
    'Refresh to make sure current record is found. An error is thrown sometimes if not
    info.Refresh
    
    'check bookmark against record count
    If bookmark > info.Recordset.RecordCount Then
        MsgBox "Could not delete record. Invalid bookmark. Please try again."
        Exit Sub
    End If
    'return to recordd to delete
    info.Recordset.bookmark = bookmark
    info.Recordset.Delete
    
    'move off empty record
    'first check to see if there are any more records
    If info.Recordset.BOF = True And info.Recordset.EOF = True Then Exit Sub
    info.Recordset.MoveNext
    If info.Recordset.EOF = True And info.Recordset.BOF <> True Then
        info.Recordset.MovePrevious
    End If
    Exit Sub
    
    
DeleteError:
    MsgBox "Cannot delete " & Err.Number & " " & Err.Description & " " & Err.Source
End Sub

Private Sub Delete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Delete.FontBold = True
    Add.FontBold = False
End Sub

Private Sub Email_Click()
    
    If Trim(txtEmail.Text) = "" Then Exit Sub
    
    'send email using Internet Explorer object and make it invisible
    Set IE = New InternetExplorer
    IE.Visible = False
    IE.Navigate "Mailto:" & txtEmail.Text
End Sub

Private Sub Email_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Email.FontBold = True
    Go.FontBold = False
End Sub

Private Sub Close_Click()
    Me.Hide
End Sub

Private Sub CLose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RecruiterForm.Close.FontBold = True
    Cancel.FontBold = False
End Sub

Private Sub Find_Click()
    
    Dim Field As String
    'local declaration to keep from stomping on add/cancel bookmark
    Dim bookmark As Variant
    Dim pos As Integer
    
    On Error Resume Next
    
    If Trim(txtFind) = "" Then Exit Sub
    
    'Find out what field to search on from combo box
    Select Case SearchField.ListIndex
        Case 0
            Field = "CompanyName"
        Case 1
            Field = "Contact"
        Case 2
            Field = "Phone"
        Case 3
            Field = "Email"
        Case 4
            Field = "Website"
        Case 5
            Field = "Address"
        Case 6
            Field = "City"
        Case 7
            Field = "State"
        Case 8
            Field = "Zip"
        Case 9
            Field = "Rating"
        Case 10
            Field = "Comments"
        Case 11
            Field = "StillOpen"
        Case 12
            Field = "JobDescription"
        Case 13
            Field = "FirstDate"
        Case 14
            Field = "NextDate"
        Case 15
            Field = "Fax"
        Case 16
            Field = "ReferredBy"
        Case Else
            Field = ""
    End Select
        
    If info.Recordset.BOF Or info.Recordset.EOF Then Exit Sub
    
    'save bookmark to current record in case find is unsuccessful so we can return to it
    bookmark = info.Recordset.bookmark
    
    'find only searches in a forward direction so start from the first record
    info.Recordset.MoveFirst
    
    'This flag is used to stop pointer repositioning in the MoveComplete event
    'We need to test for end of file here which designates that no match was
    'found.
    Searching = True
    If Field = "StillOpen" Or Field = "Rating" Then
        info.Recordset.Find Field & "= '" & Trim(txtFind) & "'"
    ElseIf Field = "FirstDate" Then
        If Not IsDate(Trim(txtFind.Text)) Then
            Beep
            txtFind.SelStart = 0
            txtFind.SelLength = Len(txtFind.Text)
            Exit Sub
        End If
        info.Recordset.Find Field & " = #" & Trim(txtFind.Text) & "#"
    ElseIf Field = "NextDate" Then
        If Not IsDate(Trim(txtFind.Text)) Then
            Beep
            txtFind.SelStart = 0
            txtFind.SelLength = Len(txtFind.Text)
            Exit Sub
        End If
        info.Recordset.Find Field & " = #" & Trim(txtFind.Text) & "#"
    Else
        info.Recordset.Find Field & " Like '" & Trim(txtFind) & "'"
    End If
    If info.Recordset.EOF Then
        MsgBox "No matching record found"
        info.Recordset.bookmark = bookmark
    End If
    Searching = False
    
    'enable button to continue search for other matching records
    FindNext.Enabled = True
    
End Sub

Private Sub Find_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Find.FontBold = True
    FindNext.FontBold = False
End Sub

Private Sub FindNext_Click()
    Dim Field As String
    Dim bookmark As Variant
    Dim response As Integer
    
    If Trim(txtFind) = "" Then
        FindNext.Enabled = False
        Exit Sub
    End If
    
    'find out what field to search from combo box
   Select Case SearchField.ListIndex
        Case 0
            Field = "CompanyName"
        Case 1
            Field = "Contact"
        Case 2
            Field = "Phone"
        Case 3
            Field = "Email"
        Case 4
            Field = "Website"
        Case 5
            Field = "Address"
        Case 6
            Field = "City"
        Case 7
            Field = "State"
        Case 8
            Field = "Zip"
        Case 9
            Field = "Rating"
        Case 10
            Field = "Comments"
        Case 11
            Field = "StillOpen"
        Case 12
            Field = "JobDescription"
        Case 13
            Field = "FirstDate"
        Case 14
            Field = "NextDate"
        Case 15
            Field = "Fax"
        Case 16
            Field = "ReferredBy"
        Case Else
            Field = ""
    End Select
    
    'save pointer in case find is unsuccessful so we can return to current record
    bookmark = info.Recordset.bookmark
    
    'flag to stop repositioning of the pointer in case of eof condition in the MoveComplete
    'event. We need to test for EOF here designating no matches
    Searching = True
    
    'move to next record we already have a match on the current one
    'otherwise it will find the current record again
    info.Recordset.MoveNext
    If info.Recordset.EOF = True Then info.Recordset.MovePrevious
    
    'On Error Resume Next
    If Field = "StillOpen" Or Field = "Rating" Then
        info.Recordset.Find Field & "= '" & Trim(txtFind) & "'"
    ElseIf Field = "FirstDate" Then
        If Not IsDate(Trim(txtFind.Text)) Then
            Beep
            txtFind.SelStart = 0
            txtFind.SelLength = Len(txtFind.Text)
            Exit Sub
        End If
        info.Recordset.Find Field & " = #" & Trim(txtFind.Text) & "#"
    ElseIf Field = "NextDate" Then
        If Not IsDate(Trim(txtFind.Text)) Then
            Beep
            txtFind.SelStart = 0
            txtFind.SelLength = Len(txtFind.Text)
            Exit Sub
        End If
        info.Recordset.Find Field & " = #" & Trim(txtFind.Text) & "#"
    Else
        info.Recordset.Find Field & " Like '" & Trim(txtFind) & "'"
    End If
    'On Error GoTo 0
    'if eof then check to see if user wants to start from the beginning
    'if not then repostion pointerto record position before search began
    If info.Recordset.EOF Then
        info.Recordset.bookmark = bookmark
        response = MsgBox("No more matching records found." & vbNewLine _
                    & "Would you like to search from the beginning again?" _
                    , vbYesNo + vbInformation, "No more matches")
        If response = vbYes Then
            Call Find_Click
            FindNext.Enabled = True
        Else
            FindNext.Enabled = False
        End If
    End If
    Searching = False
    
End Sub

Private Sub FindNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FindNext.FontBold = True
    Find.FontBold = False
    Wildcards.FontBold = False
End Sub

Private Sub FirstDate_LostFocus()
    
    'make sure date is valid
    If Not IsDate(FirstDate.Text) Then
        Beep
        FirstDate.Text = ""
    End If
    
End Sub

Private Sub Form_Activate()
   'set default state to FL change this to whatever you like
   If cboState.ListIndex = -1 Then
        cboState.ListIndex = 9 'cbostate.ListCount
    End If
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    Set DE1 = DE
    DE1.States
    Set rs = DE1.rsStates
       
    'center form
    CenterForm Me
    
    'close splash form and load combos only once
    If Started = False Then
        Unload frmSplash
        Started = True

        'load field names in searchfield box and set default values
        With SearchField
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
            .ListIndex = 11
            txtFind.Text = "true"
        End With
            
        rs.Requery
        Do While rs.EOF <> True
            cboState.AddItem rs.Fields("States")
            rs.MoveNext
        Loop
        
        For i = 1 To 5
            Rating.AddItem i
        Next
    End If
   

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'allow dragging of the form by pressing down the left mouse button and dragging
    FormDrag Me
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim ctrl As Control
    
    For Each ctrl In Controls
        If TypeOf ctrl Is CommandButton Then
            ctrl.FontBold = False
        End If
    Next
        
            
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormMDIForm Then
        Cancel = 0
    Else
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    
'    'Force window to specific size if user tries to resize.
'    'resizing throws an error if minimized so get out if minimized.
'    If Me.WindowState = vbMinimized Or Me.WindowState = vbMaximized Then Exit Sub
'    Me.Width = 9810   '6090
'    Me.Height = 5055   '6885
End Sub

Public Sub FormDrag(TheForm As Form)

    'function to allow the form to be dragged by the mousedown event
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub


Private Sub Go_Click()

    If Trim(txtWebsite.Text) = "" Then Exit Sub
    
    'make sure the address has a '.' in the address
    If Mid(txtWebsite.Text, Len(Trim(txtWebsite.Text)) - 3, 1) <> "." Then
         If Mid(txtWebsite.Text, Len(Trim(txtWebsite.Text)) - 2, 1) <> "." Then Exit Sub
    End If
    
    On Error GoTo WebError
    'create a new browser and go to website
    Set IE2 = New InternetExplorer
    IE2.Visible = True
    IE2.Navigate txtWebsite.Text
    Exit Sub
    
WebError:
    MsgBox "Cannot open Internet Explorer " & Err.Number & " " & Err.Description
End Sub

Private Sub Go_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Go.FontBold = True
    Email.FontBold = False
End Sub

Private Sub IE_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    
    'Clean up
    Set IE = Nothing
    
End Sub

Private Sub IE2_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    
    'Clean up
    Set IE2 = Nothing
    
End Sub

Private Sub info_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, _
                ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    
    'display error in my own message box and suppress the message sent by the data control
    'by settin fcanceldisplay to true and reseting errorNumber
    '3662 is a pain that keeps occurring after the form is loaded and the first rs move is made.
    'there is nothing wrong but there error gets  thrown.
    If ErrorNumber = 3662 Then
        ErrorNumber = 0
        fCancelDisplay = True
        Exit Sub
    End If
    
    'can't save empty record
    If ErrorNumber = 16389 Then
        MsgBox "You must fill in at least one field to save a record"
        On Error Resume Next
        info.Recordset.CancelUpdate
         fCancelDisplay = True
        Exit Sub
    End If
    
    If ErrorNumber <> 0 Then
        MsgBox "Trapped Error: " & ErrorNumber & " " & Description
        ErrorNumber = 0
        fCancelDisplay = True
    End If
    
End Sub

Private Sub info_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    Dim Pointer As Integer
    
    
   'test for eof or bof and move to a valid record if true
   'if we are searching skip the repositioning it will be handled in
   'the search routine. we need to know of the eof condition in that
   'routine
   
    'first check to see if there are any records
   If info.Recordset.BOF = True And info.Recordset.EOF = True Then
       RecruiterForm.Caption = "Job Tracker" & Space(53) & "0 records"
       Mainform.StatusBar.Panels(3).Text = "0 records"
       Exit Sub
    End If
   
   If info.Recordset.BOF = True Then
        Beep
        info.Recordset.MoveNext
    ElseIf info.Recordset.EOF = True Then
        If Searching = False Then
            Beep
            info.Recordset.MovePrevious
        Else
            'if searching is true don't reposition
            Exit Sub
        End If
    End If
    
    'if save and cancel buttons are enabled disable them
    Save.Enabled = False
    Cancel.Enabled = False
   
    'set record position and recordcount in caption of form
    On Error Resume Next
    Pointer = info.Recordset.AbsolutePosition
    If Pointer < 0 Then Pointer = 0
    RecruiterForm.Caption = "Job Tracker" & Space(53) & Pointer & " of " _
                & info.Recordset.RecordCount & " records"
    Mainform.StatusBar.Panels(3).Text = Pointer & " of " _
                & info.Recordset.RecordCount & " records"
    
    
    
    'Check Stillopen and turn text red if the job is still open
    If info.Recordset.BOF = True And info.Recordset.EOF = True Then Exit Sub
    If info.Recordset!StillOpen = True Then
        txtComments.ForeColor = vbRed
    Else
        txtComments.ForeColor = vbBlack
    End If
   On Error GoTo 0
End Sub

Private Sub mnuBuildQuery_Click()
    QueryForm.Show
End Sub

Private Sub Label4_Click()

    'I sized the check box to just the size of the checkbox without a label
    'so the background would show through
    'now I mimic the label of the check box with this label
    If StillOpen.Value = 1 Then
        StillOpen.Value = 0
    Else
        StillOpen.Value = 1
    End If
End Sub

Private Sub NextDate_LostFocus()

    'make sure date is valid
    If Not IsDate(NextDate.Text) Then
        Beep
        NextDate.Text = ""
    End If
    
End Sub

Private Sub Save_Click()
    
    Dim Contact As String
    Dim bookmark As Variant
    
    'move position to force a save
    On Error GoTo SaveError
    If info.Recordset.AbsolutePosition > 0 Then
        info.Recordset.MovePrevious
        info.Recordset.MoveNext
    Else
        info.Recordset.Update
        info.Refresh
    End If
    
    'disable and enable buttons
    Save.Enabled = False
    Cancel.Enabled = False
    Exit Sub
    
SaveError:
    'info.Recordset.Status
    If Err.Number = -2147467259 Then
        MsgBox "Unable to save. You must fill in at least 1 field to save a record"
        info.Recordset.CancelUpdate
        Exit Sub
    End If
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source & _
        " Unable to save record."
End Sub

Private Sub Save_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Save.FontBold = True
    Delete.FontBold = False
    Cancel.FontBold = False
End Sub

Private Sub SearchField_Change()
    txtFind.Text = ""
End Sub

Private Sub SearchField_Click()
    txtFind.Text = ""
End Sub

Private Sub StillOpen_Click()

    'if the job is still open then set the comments text to red so the record
    'is easy to identify as open
    If StillOpen.Value = 1 Then
        txtComments.ForeColor = vbRed
    Else
        StillOpen.Value = 0
        txtComments.ForeColor = vbBlack
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    
    'if user hits enter start search
    If KeyAscii = vbKeyReturn Then Call Find_Click
    
End Sub

Private Sub txtWebsite_KeyPress(KeyAscii As Integer)
    
    'if the user hits enter go to website
    If KeyAscii = vbKeyReturn Then Call Go_Click
    
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
