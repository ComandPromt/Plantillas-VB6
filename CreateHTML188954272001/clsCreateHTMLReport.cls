VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreateHTMLReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarTitle As String 'local copy
Private mvarReportHeading As String 'local copy
Private mvarSQLStatement As String 'local copy
'local variable(s) to hold property value(s)
Private mvarReportName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarValidConnection As New ADODB.Connection 'local copy
'local variable(s) to hold property value(s)
Private mvarColumnAlignment As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarBold As Boolean 'local copy
Private mvarItalic As Boolean 'local copy
Public Property Let Italic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Italic = 5
    mvarItalic = vData
End Property


Public Property Get Italic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Italic
    Italic = mvarItalic
End Property



Public Property Let Bold(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Bold = 5
    mvarBold = vData
End Property


Public Property Get Bold() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bold
    Bold = mvarBold
End Property



Public Property Let ColumnAlignment(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColumnAlignment = 5
    mvarColumnAlignment = vData
End Property


Public Property Get ColumnAlignment() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColumnAlignment
    ColumnAlignment = mvarColumnAlignment
End Property



'local variable(s) to hold property value(s)
Public Property Set ValidConnection(ByVal vData As ADODB.Connection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ValidConnection = Form1
    Set mvarValidConnection = vData
End Property
Public Property Let ReportName(ByVal vData As String)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ReportName = Form1
     mvarReportName = vData
End Property
Public Property Get ReportName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReportName
     ReportName = mvarReportName
End Property
Public Function CreateReport() As Boolean

    Dim fso As New FileSystemObject
    Dim txsFile As TextStream
    Dim i As Integer
    Dim strPercent As String
    Dim strTemp As String
    Dim rsData As New ADODB.Recordset
    
    On Error GoTo CreateErr
    
    Screen.MousePointer = vbHourglass
    
    Set txsFile = fso.CreateTextFile(mvarReportName, True)
    
    rsData.Open mvarSQLStatement, mvarValidConnection
    
    With txsFile
        .WriteLine "<html>"
        .WriteLine "<head>"
        .WriteLine "<title><b>" & mvarTitle & "</b></title>"
        .WriteLine "</head>"

        .WriteLine "<body>"

        .WriteLine "<p align=""center""><font size=""5"">" & mvarReportHeading & "</font></p>"
        .WriteLine "<table border=""1"" width=""100%"">"
        .WriteLine "<tr>"
        
        strPercent = 100 / rsData.Fields.Count - 1 & "%"
        
        For i = 0 To rsData.Fields.Count - 1
            If mvarBold Then
                Select Case mvarColumnAlignment
                    Case 0
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "left" & """><b>" & rsData.Fields(i).Name & "</b></td>"
                    Case 1
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "center" & """><b>" & rsData.Fields(i).Name & "</b></td>"
                    Case 2
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "right" & """><b>" & rsData.Fields(i).Name & "</b></td>"
                    Case Else
                        .WriteLine "<td width=""" & strPercent & """><b>" & rsData.Fields(i).Name & "</b></td>"
                End Select
                
            ElseIf mvarItalic Then
                Select Case mvarColumnAlignment
                    Case 0
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "left" & """><i>" & rsData.Fields(i).Name & "</i></td>"
                    Case 1
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "center" & """><i>" & rsData.Fields(i).Name & "</i></td>"
                    Case 2
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "right" & """><i>" & rsData.Fields(i).Name & "</i></td>"
                    Case Else
                        .WriteLine "<td width=""" & strPercent & """><i>" & rsData.Fields(i).Name & "</i></td>"
                End Select
            ElseIf mvarItalic And mvarBold Then
                Select Case mvarColumnAlignment
                    Case 0
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "left" & """><b><i>" & rsData.Fields(i).Name & "</i></b></td>"
                    Case 1
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "center" & """><b><i>" & rsData.Fields(i).Name & "</i></b></td>"
                    Case 2
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "right" & """><b><i>" & rsData.Fields(i).Name & "</i></b></td>"
                    Case Else
                        .WriteLine "<td width=""" & strPercent & """><b><i>" & rsData.Fields(i).Name & "</i></b></td>"
                End Select
            Else
                Select Case mvarColumnAlignment
                    Case 0
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "left" & """>" & rsData.Fields(i).Name & "</td>"
                    Case 1
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "center" & """>" & rsData.Fields(i).Name & "</td>"
                    Case 2
                        .WriteLine "<td width=""" & strPercent & """><p align=""" & "right" & """>" & rsData.Fields(i).Name & "</td>"
                    Case Else
                        .WriteLine "<td width=""" & strPercent & """><b>" & rsData.Fields(i).Name & "</td>"
                End Select
            End If
        Next
          .WriteLine "</tr>"
          
        Do Until rsData.EOF
            .WriteLine "<tr>"
            For i = 0 To rsData.Fields.Count - 1
                If i = 0 Then
                    If rsData(i).Value <> strTemp Then
                        .WriteLine "<td width=""" & strPercent & """>" & Format(rsData(i).Value) & "</td>"
                    Else
                        .WriteLine "<td width=""" & strPercent & """>" & Space(1) & "</td>"
                    End If
                Else
                    .WriteLine "<td width=""" & strPercent & """>" & Format(rsData(i).Value) & "</td>"
                End If
            Next i
            strTemp = rsData(0).Value
            rsData.MoveNext
              .WriteLine "</tr>"
        Loop
'
        .WriteLine "</table>"
        
        .WriteLine "</body>"
        .WriteLine "</html>"


        .Close
    End With
    Screen.MousePointer = vbNormal
    CreateReport = True
    Exit Function
CreateErr:
    CreateReport = False
    Select Case Err.Number
    
        Case Else
            MsgBox "Error #: " & Err.Number & vbCrLf & "Description: " & Err.Description
    End Select
End Function

Public Property Let SQLStatement(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SQLStatement = 5
    mvarSQLStatement = vData
End Property


Public Property Get SQLStatement() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SQLStatement
    SQLStatement = mvarSQLStatement
End Property



Public Property Let ReportHeading(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReportHeading = 5
    mvarReportHeading = vData
End Property


Public Property Get ReportHeading() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReportHeading
    ReportHeading = mvarReportHeading
End Property



Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property


Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = mvarTitle
End Property
