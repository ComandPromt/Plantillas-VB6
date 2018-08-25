Attribute VB_Name = "reportingCode"
Option Explicit
 
Sub getTimeReport(datStart As Date, datEnd As Date)
  Dim rstReport As New ADODB.Recordset
  Dim appExcel As Excel.Application
  Dim wbkReport As Excel.Workbook
  Dim wksReport As Excel.Worksheet
  Dim rngTmp As Excel.Range
  Dim lngLastID As Long
  Dim intStartRow As Integer
  Dim intLoop As Integer
  Dim strLastCustomer As String
  
  On Error GoTo errorHandler
  
  frmReport.lblProgress.Caption = "Querying Database......"
  
  cmdTimeReport.Parameters(0) = datStart
  cmdTimeReport.Parameters(1) = datEnd
  Set rstReport = returnRS(cmdTimeReport)
  If rstReport.EOF <> True Then
    rstReport.MoveFirst
    
    frmReport.lblProgress.Caption = "Building Excel Worksheet......"

    
    Set appExcel = New Excel.Application
    appExcel.Visible = True
    Set wbkReport = appExcel.Workbooks.Add
    'wbkReportame = "Kilometer Report"
    Set wksReport = wbkReport.Worksheets(1)
    wksReport.Cells(1, 1) = "Invoicing Report: Between " & datStart & " and " & datEnd
    wksReport.Cells(2, 1) = "Customer"
    wksReport.Cells(2, 2) = "Project Number"
    wksReport.Cells(2, 3) = "Project Name"
    wksReport.Cells(2, 4) = "Name"
    wksReport.Cells(2, 5) = "Hours"
    
    wksReport.Columns("A:A").ColumnWidth = 16
    Set rngTmp = wksReport.Cells(1, 1)
    rngTmp.Font.Bold = True
    rngTmp.Font.Size = 12
    Set rngTmp = wksReport.Rows(2)
    rngTmp.Font.Bold = True
    intLoop = 3
    intStartRow = intLoop
    strLastCustomer = "InitialValue"
    While rstReport.EOF <> True
      '
      'Check to see if the customer has changed
      '
      If strLastCustomer <> rstReport![strCustomer] Then
        Set rngTmp = wksReport.Rows(intLoop)
        rngTmp.Insert xlDown
        intLoop = intLoop + 1
        If Len(rstReport![strCustomer]) > 0 Then
          wksReport.Cells(intLoop, 1) = rstReport![strCustomer]
        Else
          wksReport.Cells(intLoop, 1) = "No Customer"
        End If
        strLastCustomer = rstReport![strCustomer]
      End If
      '
      'Check to see if the project has changed
      '
      If lngLastID <> rstReport![lngProjectID] Then
        wksReport.Cells(intLoop, 2) = rstReport![strProjectNumber]
        wksReport.Cells(intLoop, 3) = rstReport![strProjectName]
        lngLastID = rstReport![lngProjectID]
      End If
     
      wksReport.Cells(intLoop, 4) = rstReport![strFirstName] & " " & rstReport![strLastName]
      wksReport.Cells(intLoop, 5) = rstReport![SumOfdblHours]
      intLoop = intLoop + 1
      rstReport.MoveNext
    Wend
    
    wksReport.Columns("B:B").EntireColumn.AutoFit
    wksReport.Columns("C:C").EntireColumn.AutoFit
    wksReport.Columns("D:D").EntireColumn.AutoFit
    wksReport.Columns("e:e").EntireColumn.AutoFit
    wksReport.Columns("f:f").EntireColumn.AutoFit
    Set appExcel = Nothing
    frmReport.lblProgress.Caption = "Report Completed"
  
  Else
    MsgBox "No non-invoiced hours in the selected date range"
    frmReport.lblProgress.Caption = "No Records Returned"
  End If
  rstReport.Close
  Set rstReport = Nothing
  Exit Sub
errorHandler:
  MsgBox "#reportingCode#getTimeReport Err: " & Err.Number & " " & Err.Description
  rstReport.Close
  Set rstReport = Nothing
  Err.clear
End Sub

Sub markRangeInvoiced(datStart As Date, datEnd As Date)
  Dim rstReport As New ADODB.Recordset
  Dim intLoop As Integer
  
  cmdSelectOutstandingProjectTimes.Parameters(0) = datStart
  cmdSelectOutstandingProjectTimes.Parameters(1) = datEnd
  Set rstReport = returnRS(cmdSelectOutstandingProjectTimes)
  If rstReport.EOF <> True Then
    rstReport.MoveFirst
    While rstReport.EOF <> True
      rstReport![blnInvoiced] = True
      rstReport.Update
      frmReport.lblProgress.Caption = "Time Record ID= " & rstReport![lngID] & " updated"
      intLoop = intLoop + 1
      rstReport.MoveNext
    Wend
    frmReport.lblProgress.Caption = intLoop & " records updated"
  Else
    frmReport.lblProgress.Caption = "No Records Returned"
  End If
  rstReport.Close
  Set rstReport = Nothing
End Sub

