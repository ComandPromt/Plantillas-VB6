VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStudent 
   BorderStyle     =   0  'None
   Caption         =   "Computerized Testing Program - Student Screen"
   ClientHeight    =   8220
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStudent.frx":0000
   ScaleHeight     =   8220
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   6960
   End
   Begin VB.Data datScores 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Data datLogin 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2100
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   4320
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFinished 
      Caption         =   "Finished"
      Height          =   495
      Left            =   9480
      TabIndex        =   14
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous Question"
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdNextQuestion 
      Caption         =   "Next Question"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Frame frmAnswers 
      Caption         =   "Answers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1200
      TabIndex        =   5
      Top             =   4920
      Width           =   9615
      Begin VB.OptionButton opt4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   1680
         Width           =   7575
      End
      Begin VB.OptionButton opt3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1200
         Width           =   7575
      End
      Begin VB.OptionButton opt2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   7575
      End
      Begin VB.OptionButton opt1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.TextBox txtQuestion 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label lblTestName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label lblQuestionType 
      BackStyle       =   0  'Transparent
      Caption         =   "Question Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Question #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -120
      TabIndex        =   11
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblQuestionNumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblSeconds 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblColon 
      BackColor       =   &H8000000E&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblMinute 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7920
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   7920
      Shape           =   2  'Oval
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Test"
      End
      Begin VB.Menu mnuHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "&Start Test"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuStudent 
      Caption         =   "S&tudent Information"
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuHyphen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisplayScores 
         Caption         =   "&Display Test Scores"
      End
      Begin VB.Menu mnuPrintScores 
         Caption         =   "&Print Test Scores"
      End
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentQuest As Integer
Dim questNum As Integer
Dim goBack As Boolean
Dim timed As Boolean
Dim minutes As Integer


Private Sub cmdFinished_Click()

    'check for no answer given
    If opt1.Value = False And opt2.Value = False And opt3.Value = False _
       And opt4.Value = False Then
       MsgBox "You must select an answer!", , "Try Again!"
       Exit Sub
    End If
    
    'assign answer selected to users answer array for checking
    If opt1.Value = True Then
        usersAnswer(currentQuest) = "A"
    Else
        If opt2.Value = True Then
            usersAnswer(currentQuest) = "B"
        Else
            If opt3.Value = True Then
                usersAnswer(currentQuest) = "C"
            Else
                usersAnswer(currentQuest) = "D"
            End If
        End If
    End If
    
    testIsOver
        
End Sub

Private Sub cmdNextQuestion_Click()

    'check for no answer
    If opt1.Value = False And opt2.Value = False And opt3.Value = False _
       And opt4.Value = False Then
       MsgBox "You must select an answer!", , "Try Again!"
       Exit Sub
    End If
    
    'enable previous command button if go back is allowed
    If goBack Then
        cmdPrevious.Enabled = True
    End If
    
    'assigns answer to users answer array
    If opt1.Value = True Then
        usersAnswer(currentQuest) = "A"
    Else
        If opt2.Value = True Then
            usersAnswer(currentQuest) = "B"
        Else
            If opt3.Value = True Then
                usersAnswer(currentQuest) = "C"
            Else
                usersAnswer(currentQuest) = "D"
            End If
        End If
    End If
    
    currentQuest = currentQuest + 1
    
    'if current question is last question then disable next button
    If currentQuest = questNum Then
        cmdNextQuestion.Enabled = False
        cmdFinished.Enabled = True
    End If
    
    'load value (if any) from users answer array for next question
    If usersAnswer(currentQuest) = "A" Then
        opt1.Value = True
    Else
        If usersAnswer(currentQuest) = "B" Then
            opt2.Value = True
        Else
            If usersAnswer(currentQuest) = "C" Then
                opt3.Value = True
            Else
                If usersAnswer(currentQuest) = "D" Then
                    opt4.Value = True
                Else
                    opt1.Value = False
                    opt2.Value = False
                    opt3.Value = False
                    opt4.Value = False
                End If
            End If
        End If
    End If
    
    'load question onto screen for user viewing
    loadQuestion
    
End Sub

Private Sub cmdPrevious_Click()

    
    'make sure question is not number one when moving back
    If currentQuest > 1 Then
    
        'assign current answer to users answer array
        If opt1.Value = True Then
            usersAnswer(currentQuest) = "A"
        Else
            If opt2.Value = True Then
                usersAnswer(currentQuest) = "B"
            Else
                If opt3.Value = True Then
                    usersAnswer(currentQuest) = "C"
                Else
                    If opt4.Value = True Then
                        usersAnswer(currentQuest) = "D"
                    End If
                End If
            End If
        End If
            
        'enable next button
        cmdNextQuestion.Enabled = True
        currentQuest = currentQuest - 1
        
        'if current question is number 1 disable previous button
        If currentQuest = 1 Then
            cmdPrevious.Enabled = False
        End If
        'load current answer(if any) from users answer array for the
        'question the user is moving back to
        If usersAnswer(currentQuest) = "A" Then
            opt1.Value = True
        Else
            If usersAnswer(currentQuest) = "B" Then
                opt2.Value = True
            Else
                If usersAnswer(currentQuest) = "C" Then
                    opt3.Value = True
                Else
                    If usersAnswer(currentQuest) = "D" Then
                        opt4.Value = True
                    Else
                        opt1.Value = False
                        opt2.Value = False
                        opt3.Value = False
                        opt4.Value = False
                    End If
                End If
            End If
        End If
        loadQuestion
    End If
    
End Sub

Private Sub Form_Load()

    Show
    Unload frmLogin
    Caption = "Computerized Testing Program - Student Screen - User - " & loggedUser
    opt1.Value = False
    cmdPrevious.Enabled = False
    cmdNextQuestion.Enabled = False
    cmdFinished.Enabled = False
    
    datLogin.DatabaseName = App.Path & "\login.mdb"
    datLogin.RecordSource = "Login"
    datLogin.Refresh
    
    datScores.DatabaseName = App.Path & "\login.mdb"
    datScores.RecordSource = "TestScores"
    datScores.Refresh
    
End Sub


Private Sub mnuChangePassword_Click()

    'edit currently logged on user's password
    Load frmTeachPass
    frmTeachPass.Caption = "Change Student Password"
    'search for user in database and display current(old) password
    With datLogin.Recordset
        .MoveFirst
        Do Until .EOF
            If loggedUser = .Fields("UserID").Value Then
                frmTeachPass.picOutput.Cls
                frmTeachPass.picOutput.Print "Old Password:  "; .Fields("Password").Value
                Exit Do
            End If
            .MoveNext
        Loop
    End With
                
End Sub

Private Sub mnuDisplayScores_Click()

    Load frmDisplayTestStudent
    
End Sub

Private Sub mnuExit_Click()

    End
    
End Sub

Private Sub mnuOpen_Click()

    Dim i As Integer
    Dim isTimed As String
    Dim allowBack As String
    Dim theColor As String
    Dim ColorSet As Long
        
    On Error GoTo dlgError
    
    With dlgFile
        .CancelError = True
        .Flags = 2
        .Filter = "Test Files| *.tst"
        .DialogTitle = "Choose Test To Take"
        .ShowOpen
            
        'clear test question array
        For i = 1 To 100
            usersAnswer(i) = ""
            questTest(i).answerA = ""
            questTest(i).answerB = ""
            questTest(i).answerC = ""
            questTest(i).answerD = ""
            questTest(i).correctAns = ""
            questTest(i).quest = ""
            questTest(i).theType = ""
        Next i
        
        'set current test for the program to reference
        currentTest = .FileName
        'set the label to indicate current test
        lblTestName.Caption = currentTest
        
        questNum = 0
        'load test question array from test file
        Open currentTest For Input As #1
        Do Until EOF(1)
            questNum = questNum + 1
            Input #1, questTest(questNum).quest
            Input #1, questTest(questNum).answerA
            Input #1, questTest(questNum).answerB
            Input #1, questTest(questNum).answerC
            Input #1, questTest(questNum).answerD
            Input #1, questTest(questNum).correctAns
            Input #1, questTest(questNum).theType
        Loop
        Close #1
                
        'load layout for test
        Open Left(currentTest, Len(currentTest) - 3) & "lyt" For Input As #1
            Input #1, isTimed
            Input #1, minutes
            Input #1, allowBack
            Input #1, theColor
        Close #1
            
            If isTimed = "T" Then
                timed = True
            Else
                timed = False
            End If
            
            If timed Then
                lblMinute.Caption = minutes
            Else
                lblMinute.Caption = 0
            End If
            lblSeconds.Caption = "00"
            Timer1.Enabled = False
            
            If allowBack = "T" Then
                goBack = True
            Else
                goBack = False
            End If
                                    
            cmdPrevious.Enabled = False
            
            If theColor = "R" Then
                ColorSet = vbRed
            Else
                If theColor = "GN" Then
                    ColorSet = vbGreen
                Else
                    If theColor = "B" Then
                        ColorSet = vbBlue
                    Else
                        ColorSet = -2147483633
                    End If
                End If
            End If
            
            frmStudent.BackColor = ColorSet
            Label2.BackColor = ColorSet
            lblQuestionNumber.BackColor = ColorSet
            lblTestName.BackColor = ColorSet
            lblQuestionType.BackColor = ColorSet
            frmAnswers.BackColor = ColorSet
            opt1.BackColor = ColorSet
            opt2.BackColor = ColorSet
            opt3.BackColor = ColorSet
            opt4.BackColor = ColorSet
            cmdPrevious.BackColor = ColorSet
            cmdNextQuestion.BackColor = ColorSet
            cmdFinished.BackColor = ColorSet
            
        mnuStart.Enabled = True
    End With
               
dlgError:
    On Error GoTo 0
    Exit Sub
    
End Sub

Private Sub mnuPrintScores_Click()
    
    'set printer properties
    Printer.ScaleMode = 4
    Printer.FontSize = 12
    Printer.CurrentY = 5
    
    'print all test scores for currently logged user
    With datScores.Recordset
        .MoveFirst
        Do Until .EOF
            If .Fields("ID").Value = loggedUser Then
                Printer.CurrentX = 5
                Printer.Print .Fields("Test").Value;
                Printer.CurrentX = 25
                Printer.Print .Fields("Date").Value;
                Printer.CurrentX = 45
                Printer.Print .Fields("Grade").Value
            End If
            .MoveNext
        Loop
    End With
 
    Printer.EndDoc
    
End Sub

Private Sub mnuStart_Click()

    Dim response As Integer
    
    'check for timed test, and give user option to start or not
    If timed Then
        response = MsgBox("You have " & minutes & " minutes to finish this test.  " & _
               "Once you have started you can't stop the test.  " & _
               "Are you ready to start?", vbYesNo, "Timed Test")
    Else
        response = MsgBox("You have an unlimited amount of time to finish this " & _
                "test.  Once you have started you can't stop the " & _
                "test.  Are you ready to start?", vbYesNo, _
                "Non-Timed Test")
    End If
    
    If response = vbNo Then
        Exit Sub
    Else
        If timed Then
            Timer1.Enabled = True
        End If
        cmdNextQuestion.Enabled = True
        cmdFinished.Enabled = False
        currentQuest = 1
        mnuOpen.Enabled = False
        mnuStart.Enabled = False
        mnuExit.Enabled = False
        mnuChangePassword.Enabled = False
        mnuDisplayScores.Enabled = False
        mnuPrintScores.Enabled = False
        
        loadQuestion
    End If
                
End Sub

Private Sub loadQuestion()

    'set the label to display current question number
    lblQuestionNumber.Caption = currentQuest
    
    'load answers and question on to user screen
    opt1.Caption = questTest(currentQuest).answerA
    opt2.Caption = questTest(currentQuest).answerB
    opt3.Caption = questTest(currentQuest).answerC
    opt4.Caption = questTest(currentQuest).answerD
    txtQuestion.Text = questTest(currentQuest).quest
            
    'if there is only 1 question in test disable all buttons except
    'for finished.
    If questNum = 1 Then
        cmdNextQuestion.Enabled = False
        cmdFinished.Enabled = True
        cmdPrevious.Enabled = False
    End If
       
    'set the label to indicate the type of question
    If questTest(currentQuest).theType = "T" Then
        opt3.Enabled = False
        opt4.Enabled = False
        lblQuestionType.Caption = "True/False"
    Else
        opt3.Enabled = True
        opt4.Enabled = True
        lblQuestionType.Caption = "Multiple Choice"
    End If
    
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()

    'stop test if no time is left
    If Val(lblMinute.Caption) = 0 And Val(lblSeconds.Caption) = 0 Then
        testIsOver
        Exit Sub
    End If
    
    'decrement timer by 1 second every 1 second
    If Val(lblSeconds.Caption) = 0 Then
        lblSeconds.Caption = 59
        lblMinute.Caption = Val(lblMinute.Caption) - 1
    Else
        lblSeconds.Caption = Val(lblSeconds.Caption) - 1
        If Val(lblSeconds.Caption) < 10 Then
            lblSeconds.Caption = "0" & lblSeconds.Caption
        End If
    End If
    
End Sub

Private Sub testIsOver()

    Dim i As Integer
    Dim score As Integer
    Dim numOfQuestions As Integer
    Dim numberCorrect As Integer
    
    'shut off timer, and display test over
    Timer1.Enabled = False
    MsgBox "Time is up!!", , "Test Over"
    'enable menu options
    mnuOpen.Enabled = True
    mnuExit.Enabled = True
    mnuChangePassword.Enabled = True
    mnuDisplayScores.Enabled = True
    mnuPrintScores.Enabled = True
    'disable command buttons
    cmdPrevious.Enabled = False
    cmdNextQuestion.Enabled = False
    cmdFinished.Enabled = False
    
    'check answers
    numberCorrect = 0
    numOfQuestions = questNum
    For i = 1 To numOfQuestions
        If usersAnswer(i) = RTrim(questTest(i).correctAns) Then
            numberCorrect = numberCorrect + 1
        End If
    Next i
    score = numberCorrect / questNum * 100
    userScore = score
    numOfQ = questNum
    numCorrect = numberCorrect
    
    'add test, id, date, and grade to test scores database
    With datScores.Recordset
        .AddNew
        .Fields("Test").Value = Mid(currentTest, 1, 50)
        .Fields("ID").Value = loggedUser
        .Fields("Date").Value = Date
        .Fields("Grade").Value = score
        .Update
    End With
    
    Load frmDisplayGrade
    
End Sub


