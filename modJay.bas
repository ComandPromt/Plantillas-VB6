Attribute VB_Name = "modJay"
Option Explicit
'Walther Musch: example made 6 januari 1998
'(c) KATHER Produkties 1998
'use with permission to alter
'some suggestions:
'   save after Calculate the result on file/database cq disk
'   print option
'   the Function CheckTime can be improved
'   in use you can set the property Visible = False for the controls called Label7(#)

Public Sub SetDate(richting%, Optional StartDate)
'get startingdate and place labels

    Dim Vandaag As Date
    Dim vandaagweekdag As Integer
    Dim t As Integer
    
    If IsMissing(StartDate) Then
        Vandaag = Format(Now, "dd-mm-yyyy")
    Else
        Vandaag = Format(StartDate, "dd-mm-yyyy")
    End If
    
    Select Case DatePart("w", Vandaag)
    Case 1
        vandaagweekdag% = 7
    Case 2
        vandaagweekdag% = 1
    Case 3
        vandaagweekdag% = 2
    Case 4
        vandaagweekdag% = 3
    Case 5
        vandaagweekdag% = 4
    Case 6
        vandaagweekdag% = 5
    Case 7
        vandaagweekdag% = 6
    End Select
    
    Select Case richting%
    Case -1    'voorgaande week
        Vandaag = DateAdd("d", -7, frmMain.Label1(vandaagweekdag))
    Case 1    'volgende week
        Vandaag = DateAdd("d", 7, frmMain.Label1(vandaagweekdag))
    End Select
    
    vandaagweekdag% = DatePart("w", Vandaag)
    'calculate first day
    Select Case vandaagweekdag%
    Case 2    'monday
        Vandaag = DateAdd("d", -1, Vandaag)
    Case 3    'thuesday
        Vandaag = DateAdd("d", -2, Vandaag)
    Case 4    'wensday
        Vandaag = DateAdd("d", -3, Vandaag)
    Case 5    'thurstday
        Vandaag = DateAdd("d", -4, Vandaag)
    Case 6    'friday
        Vandaag = DateAdd("d", -5, Vandaag)
    Case 7    'saterday
        Vandaag = DateAdd("d", -6, Vandaag)
    Case 1    'sunday
        Vandaag = DateAdd("d", -7, Vandaag)
    End Select
    
    For t% = 1 To 7
        Vandaag = DateAdd("d", 1, Vandaag)
        frmMain.Label1(t%) = Format(Vandaag, "dd-mm-yyyy")
        frmMain.Label2(t%) = WhichDay(DatePart("w", Vandaag))
    Next t%
    
End Sub
Public Sub ClearInputBox()
'clear inputboxes and some labels
    Dim bX As Byte
    
    On Error Resume Next
    For bX = 0 To 13
        If bX > 0 And bX < 8 Then
            frmMain.Label5(bX).Caption = ""
            frmMain.Label7(bX).Caption = ""
        End If
        frmMain.Text1(bX).Text = ""
    Next bX
    frmMain.Text1(0).SetFocus
    frmMain.Label7(8).Caption = ""
    frmMain.Label5(8).Caption = ""
End Sub
Private Function WhichDay(Source As Integer) As String
'function to show weekday name
    Select Case Source
    Case 1
        WhichDay = "Sunday"
    Case 2
        WhichDay = "Monday"
    Case 3
        WhichDay = "Thuesday"
    Case 4
        WhichDay = "Wednsday"
    Case 5
        WhichDay = "Thurstday"
    Case 6
        WhichDay = "Friday"
    Case 7
        WhichDay = "Saterday"
    End Select
    
End Function
Public Function CheckTijd(bron$) As Date
'convert almost any input to real time
'always open for better suggestions!
    Dim t%
    'checking on digits and seperator
    Const Getal$ = "1234567890.:"
    
    For t% = 1 To Len(bron$)
        If InStr(Getal$, Mid$(bron$, t%, 1)) = 0 Then Exit Function
        If Mid$(bron$, t%, 1) = "." Then bron$ = Left$(bron$, t% - 1) & ":" & Right$(bron$, Len(bron$) - t%)
    Next t%
    
    Select Case Len(bron$)
    Case 0
        Exit Function
    Case 1
        bron$ = "0" & bron$ & ":00"
    Case 2
        bron$ = bron$ & ":00"
    Case 3
        t% = InStr(bron$, ":")
        If t% = 0 Then _
            bron$ = Left$(bron$, 1) & ":" & Right$(bron$, 2)
    Case 4
        t% = InStr(bron$, ":")
        If t% = 0 Then _
            bron$ = Left$(bron$, 2) & ":" & Right$(bron$, 2)
    Case 5
        bron$ = Left$(bron$, 2) & ":" & Right$(bron$, 2)
    End Select
    
    On Error Resume Next
    CheckTijd = TimeValue(bron$)
End Function
Public Sub SetSelected()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
   
End Sub
Private Function CalculateHours(Source As Double) As String
'function to get from DateDiff the hours en minutes
    Dim dblHours As Double
    Dim bHours As Byte
    Dim bMinutes As Byte
    Dim p
    Dim strDummy As String
    
    bMinutes = 0
    dblHours = Source / 60
    bHours = Int(dblHours)
        
    strDummy = CStr(dblHours)
    p = InStr(strDummy, ",")
    If p <> 0 Then
        strDummy = Right$(strDummy, Len(strDummy) - p)
        strDummy = Left$(strDummy, 2)
        bMinutes = CByte((60 * CByte(strDummy)) / 100)
    End If
    CalculateHours = CStr(bHours) & "." & CStr(bMinutes)
    
End Function
Public Sub GetDate()
'procedure to start with given date
    Dim dInput
    
    dInput = InputBox("Which date to start with?", , Format(Now, "Short Date"))
    If dInput = "" Then Exit Sub
    
    Call SetDate(0, dInput)
    Call ClearInputBox
    
End Sub
Public Sub CalculateWorkingHours()
'procedure to calculated working time
'get starttime and endtime
'calculate difference with DateDiff
'calculate working hours with CalculateHours
    Dim bX As Byte, bY As Byte
    Dim startTime As Date
    Dim endTime As Date
    Dim dblTime As Double
    
    On Error Resume Next
    bY = 1
    For bX = 0 To 13 Step 2
        If frmMain.Text1(bX) <> "" Then
            If frmMain.Text1(bX) <> CheckTijd("0") Then
                startTime = TimeValue(frmMain.Text1(bX).Text)
                endTime = TimeValue(frmMain.Text1(bX + 1).Text)
                frmMain.Label5(bY).Caption = CalculateHours(CDbl(DateDiff("n", startTime, endTime)))
                frmMain.Label7(bY).Caption = DateDiff("n", startTime, endTime)
            End If
        bY = bY + 1
        End If
    Next bX
    
    'total working hours
    dblTime = 0
    For bX = 1 To 7
        If frmMain.Label7(bX).Caption <> "" Then
            dblTime = dblTime + CDbl(frmMain.Label7(bX).Caption)
        End If
    Next bX
    frmMain.Label7(8).Caption = CStr(dblTime)
    frmMain.Label5(8).Caption = CalculateHours(dblTime)
    
End Sub
