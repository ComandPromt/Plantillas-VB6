Attribute VB_Name = "Module1"
Option Explicit

Private DBs As String                          ' Database path and file name
Private Database1 As Database        ' Database object
Private A1 As Object, A2 As Object  ' Database recordsets
Private DBOptions As String
Private DBLocale As String
Private DBPassword As String
Private DBs2 As String                        ' Temp database
Private DBTable As String

Public Function GotoAirline(AirlineName As String) As Boolean
    frmMain.Accounts.MoveFirst
    Do Until frmMain.Accounts!Airline = AirlineName
        If frmMain.Accounts.EOF = True Then GoTo EOFF
        frmMain.Accounts.MoveNext
    Loop
    
    frmMain.AccountInfo.MoveFirst
    frmMain.AccountInfo.Edit
    Do Until frmMain.AccountInfo!Airline = AirlineName
        If frmMain.AccountInfo.EOF = True Then GoTo EOFF
        frmMain.AccountInfo.MoveNext
    Loop

    frmMain.Details.MoveFirst
    Do Until frmMain.Details!Airline = AirlineName
        If frmMain.Details.EOF = True Then GoTo EOFF
        frmMain.Details.MoveNext
        If frmMain.Details.EOF = True Then GoTo EOFF
    Loop
    
    GoTo SkipIt
EOFF:
    GotoAirline = False
    Exit Function
SkipIt:
    GotoAirline = True
    Exit Function
End Function

Public Function LoadData(AirlineName As String)
    If GotoAirline(AirlineName) = False Then Exit Function
    frmMain.Loading = True
    
    With frmMain
        .txtBillingAddress.Text = .AccountInfo!Address
        .txtBillingName.Text = .AccountInfo!Name
        .txtCityStateZip.Text = .AccountInfo!CityStateZip
        If .AccountInfo!Other = "None" Then
            .txtOther.Text = ""
        Else
            .txtOther.Text = .AccountInfo!Other
        End If
        .txtFields(0).Text = .Accounts!Airline
        .txtFields(1).Text = .Accounts!CurrentCharges
        .txtFields(2).Text = .Accounts!InvoiceNum
        .txtFields(3).Text = .Accounts!Operations
        .txtFields(4).Text = .Accounts!OutstandingBalance
        .txtFields(5).Text = .Accounts!Rate
        .txtFields(6).Text = .Accounts!Total
'        If .Accounts!Notes = Null Then
'            .txtNotes.Text = ""
'        Else
'            .txtNotes.Text = .Accounts!Notes
'        End If
        .cboPaid.Text = .Accounts!Paid
    End With
    frmMain.Loading = False
End Function

Public Function SaveData()
    If GotoAirline(frmMain.txtFields(0).Text) = False Then Exit Function
    
    frmMain.Loading = True
    With frmMain
        .AccountInfo.Edit
        .Accounts.Edit
        .AccountInfo!Airline = .txtFields(0).Text
        .AccountInfo!Address = .txtBillingAddress.Text
        .AccountInfo!Name = .txtBillingName.Text
        .AccountInfo!CityStateZip = .txtCityStateZip.Text
        If .txtOther.Text = "" Then
            .AccountInfo!Other = "None"
        Else
            .AccountInfo!Other = .txtOther.Text
        End If
        .Accounts!Airline = .txtFields(0).Text
        .Accounts!CurrentCharges = .txtFields(1).Text
        .Accounts!InvoiceNum = .txtFields(2).Text
        .Accounts!Operations = .txtFields(3).Text
        .Accounts!OutstandingBalance = .txtFields(4).Text
        .Accounts!Rate = .txtFields(5).Text
        .Accounts!Total = .txtFields(6).Text
        .Accounts!Paid = .cboPaid.Text
        .Accounts!Notes = .txtNotes.Text
        
        
        .Accounts.Update
        .AccountInfo.Update
        
        Open App.Path & "\Charges For.txt" For Output As #1
            Print #1, frmMain.txtPC.Text
        Close #1
    End With
    frmMain.Loading = False
End Function

Public Function SaveNewData()
    Dim Temp As String
    
    frmMain.Loading = True
    With frmMain
        .AccountInfo.AddNew
        .Accounts.AddNew
        .Details.AddNew
        DoEvents
        .Details!Airline = .txtFields(0).Text
        .Details!Flight = "None"
        .Details!Airbill = "None"
        .Details!Destination = "None"
        .Details!Description = "None"
        .Details!Operations = "None"
        .Details!Date = "None"
        .Details!Weight = "None"
        .Details!POnum = "None"
        .Details!Aircraftnum = "None"
        .Details.Update
        .AccountInfo!Airline = .txtFields(0).Text
        .AccountInfo!Address = .txtBillingAddress.Text
        .AccountInfo!Name = .txtBillingName.Text
        .AccountInfo!CityStateZip = .txtCityStateZip.Text
        If .txtOther.Text = "" Then
            .AccountInfo!Other = "None"
        Else
            .AccountInfo!Other = .txtOther.Text
        End If
        .Accounts!Airline = .txtFields(0).Text
        .Accounts!CurrentCharges = .txtFields(1).Text
        .Accounts!InvoiceNum = .txtFields(2).Text
        .Accounts!Operations = .txtFields(3).Text
        .Accounts!OutstandingBalance = .txtFields(4).Text
        .Accounts!Rate = .txtFields(5).Text
        .Accounts!Total = .txtFields(6).Text
        .Accounts!Paid = .cboPaid.Text
        .Accounts!Notes = .txtNotes.Text
        
        .Op = "0"
        .Accounts.Update
        .AccountInfo.Update
        .Options.Edit
        .Options!CurrentInvoiceNum = .Options!CurrentInvoiceNum + 1
        .Options.Update
        
        Open App.Path & "\Charges For.txt" For Output As #1
            Print #1, frmMain.txtPC.Text
        Close #1
        
    End With
    frmMain.Loading = False
End Function

Public Function RefreshListbox()
    frmMain.Accounts.MoveFirst
    frmMain.Accounts.Edit
    frmMain.Loading = True
    frmMain.lstAccounts.Clear
    Do Until frmMain.Accounts.EOF = True
        frmMain.lstAccounts.AddItem frmMain.Accounts!Airline
        frmMain.Accounts.MoveNext
    Loop
    frmMain.Loading = False
End Function

Public Function LoadDetails(AirlineName As String)
    Dim Temp As String
    Dim Temp2 As String
        
    If GotoAirline(AirlineName) = False Then Exit Function
    
    With frmMain
        .Details.Edit
        
        'Flight
        Temp = .Details!Flight
        If Temp = "None" Then
            .txtFlight.Text = ""
            .txtFlight.BackColor = &H8000000F
            .txtFlight.Enabled = False
        Else
            .txtFlight.BackColor = &H80000005
            .txtFlight.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #1
                Print #1, Temp
            Close #1
            Open App.Path & "\Temp2.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtFlight.Text = Temp2
                    End If
                Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
        
        'Airbill
        Temp = .Details!Airbill
        If Temp = "None" Then
            .txtAirbill.Text = ""
            .txtAirbill.BackColor = &H8000000F
            .txtAirbill.Enabled = False
        Else
            .txtAirbill.BackColor = &H80000005
            .txtAirbill.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #1
                Print #1, Temp
            Close #1
            Open App.Path & "\Temp2.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtAirbill.Text = Temp2
                    End If
                Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
        
        'Destination
        Temp = .Details!Destination
        If Temp = "None" Then
            .txtDestination.Text = ""
            .txtDestination.BackColor = &H8000000F
            .txtDestination.Enabled = False
        Else
            .txtDestination.BackColor = &H80000005
            .txtDestination.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #1
                Print #1, Temp
            Close #1
            Open App.Path & "\Temp2.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtDestination.Text = Temp2
                    End If
                Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
                
        'Description
        Temp = .Details!Description
        If Temp = "None" Then
            .txtDescription.Text = ""
            .txtDescription.BackColor = &H8000000F
            .txtDescription.Enabled = False
        Else
            .txtDescription.BackColor = &H80000005
            .txtDescription.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #1
                Print #1, Temp
            Close #1
            Open App.Path & "\Temp2.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtDescription.Text = Temp2
                    End If
                Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
                
        'Operations
        Temp = .Details!Operations
        If Temp = "None" Then
            .txtOperations.Text = ""
            .txtOperations.BackColor = &H8000000F
            .txtOperations.Enabled = False
        Else
            .txtOperations.BackColor = &H80000005
            .txtOperations.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #1
                Print #1, Temp
            Close #1
            Open App.Path & "\Temp2.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtOperations.Text = Temp2
                    End If
                Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
                
        'Date
        Temp = .Details!Date
        If Temp = "None" Then
            .txtDate.Text = ""
            .txtDate.BackColor = &H8000000F
            .txtDate.Enabled = False
        Else
            .txtDate.BackColor = &H80000005
            .txtDate.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #2
                Print #2, Temp
            Close #2
            Open App.Path & "\Temp2.txt" For Input As #1
            Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtDate.Text = Temp2
                    End If
            Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
        
        'Weight
        Temp = .Details!Weight
        If Temp = "None" Then
            .txtWeight.Text = ""
            .txtWeight.BackColor = &H8000000F
            .txtWeight.Enabled = False
        Else
            .txtWeight.BackColor = &H80000005
            .txtWeight.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #1
                Print #1, Temp
            Close #1
            Open App.Path & "\Temp2.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtWeight.Text = Temp2
                    End If
                Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
        
        'PO Number
        Temp = .Details!POnum
        If Temp = "None" Then
            .txtPO.Text = ""
            .txtPO.BackColor = &H8000000F
            .txtPO.Enabled = False
        Else
            .txtPO.BackColor = &H80000005
            .txtPO.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #1
                Print #1, Temp
            Close #1
            Open App.Path & "\Temp2.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtPO.Text = Temp2
                    End If
                Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
        
        'Aircraft Number
        Temp = .Details!Aircraftnum
        If Temp = "None" Then
            .txtAircraft.Text = ""
            .txtAircraft.BackColor = &H8000000F
            .txtAircraft.Enabled = False
        Else
            .txtAircraft.BackColor = &H80000005
            .txtAircraft.Enabled = True
            Open App.Path & "\Temp2.txt" For Output As #1
                Print #1, Temp
            Close #1
            Open App.Path & "\Temp2.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, Temp
                    Input #1, Temp2
                    If Temp = .Op Then
                        .txtAircraft.Text = Temp2
                    End If
                Loop
            Close #1
            Kill App.Path & "\Temp2.txt"
        End If
        
    End With
    
    frmMain.lblOP.Caption = "Current Operation: " & frmMain.Op
End Function

Public Function SaveDetailData(AirlineName As String)
    If GotoAirline(AirlineName) = False Then Exit Function
    
    Dim Temp As String
    Dim Temp2 As String
    Dim Data As String
    Dim It As String
    Dim TheRest As String
    Dim strInput As String
    Dim NI As String    'NI = Not Important
    Dim IsAValue As Boolean
    
    With frmMain
        .Details.Edit
        
        'Flight
        If .txtFlight.Text <> "" Then
            Data = .Details!Flight
            If Data = "None" Then
                .Details!Flight = .Op & "," & .txtFlight.Text
                GoTo Next1
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!Flight = .Details!Flight & "," & .Op & "," & .txtFlight.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!Flight
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!Flight = Data & .Op & "," & .txtFlight.Text
                Else
                    .Details!Flight = Data & "," & .Op & "," & .txtFlight.Text
                End If
            End If
        Else
            If .Details!Flight = "None" Or .Details!Flight = "" Then
                .Details!Flight = "None"
            End If
        End If
Next1:
        
        
        'Airbill
        If .txtAirbill.Text <> "" Then
            Data = .Details!Airbill
            If Data = "None" Then
                .Details!Airbill = .Op & "," & .txtAirbill.Text
                GoTo Next2
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!Airbill = .Details!Airbill & "," & .Op & "," & .txtAirbill.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!Airbill
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt2
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt2:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!Airbill = Data & .Op & "," & .txtAirbill.Text
                Else
                    .Details!Airbill = Data & "," & .Op & "," & .txtAirbill.Text
                End If
            End If
        Else
            If .Details!Airbill = "None" Or .Details!Airbill = "" Then
                .Details!Airbill = "None"
            End If
        End If
Next2:
        
        
        'Destination
        If .txtDestination.Text <> "" Then
            Data = .Details!Destination
            If Data = "None" Then
                .Details!Destination = .Op & "," & .txtDestination.Text
                GoTo Next3
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!Destination = .Details!Destination & "," & .Op & "," & .txtDestination.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!Destination
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt3
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt3:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!Destination = Data & .Op & "," & .txtDestination.Text
                Else
                    .Details!Destination = Data & "," & .Op & "," & .txtDestination.Text
                End If
            End If
        Else
            If .Details!Destination = "None" Or .Details!Destination = "" Then
                .Details!Destination = "None"
            End If
        End If
Next3:

        'Description
        If .txtDescription.Text <> "" Then
            Data = .Details!Description
            If Data = "None" Then
                .Details!Description = .Op & "," & .txtDescription.Text
                GoTo Next4
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!Description = .Details!Description & "," & .Op & "," & .txtDescription.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!Description
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt4
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt4:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!Description = Data & .Op & "," & .txtDescription.Text
                Else
                    .Details!Description = Data & "," & .Op & "," & .txtDescription.Text
                End If
            End If
        Else
            If .Details!Description = "None" Or .Details!Description = "" Then
                .Details!Description = "None"
            End If
        End If
Next4:
        
        'Operations
        If .txtOperations.Text <> "" Then
            Data = .Details!Operations
            If Data = "None" Then
                .Details!Operations = .Op & "," & .txtOperations.Text
                GoTo Next5
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!Operations = .Details!Operations & "," & .Op & "," & .txtOperations.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!Operations
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt5
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt5:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!Operations = Data & .Op & "," & .txtOperations.Text
                Else
                    .Details!Operations = Data & "," & .Op & "," & .txtOperations.Text
                End If
            End If
        Else
            If .Details!Operations = "None" Or .Details!Operations = "" Then
                .Details!Operations = "None"
            End If
        End If
Next5:
                
        'Date
        If .txtDate.Text <> "" Then
            Data = .Details!Date
            If Data = "None" Then
                .Details!Date = .Op & "," & .txtDate.Text
                GoTo Next6
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!Date = .Details!Date & "," & .Op & "," & .txtDate.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!Date
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt6
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt6:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!Date = Data & .Op & "," & .txtDate.Text
                Else
                    .Details!Date = Data & "," & .Op & "," & .txtDate.Text
                End If
            End If
        Else
            If .Details!Date = "None" Or .Details!Date = "" Then
                .Details!Date = "None"
            End If
        End If
Next6:
        
        'Weight
        If .txtWeight.Text <> "" Then
            Data = .Details!Weight
            If Data = "None" Then
                .Details!Weight = .Op & "," & .txtWeight.Text
                GoTo Next7
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!Weight = .Details!Weight & "," & .Op & "," & .txtWeight.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!Weight
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt7
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt7:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!Weight = Data & .Op & "," & .txtWeight.Text
                Else
                    .Details!Weight = Data & "," & .Op & "," & .txtWeight.Text
                End If
            End If
        Else
            If .Details!Weight = "None" Or .Details!Weight = "" Then
                .Details!Weight = "None"
            End If
        End If
Next7:
        
        'PO Number
        If .txtPO.Text <> "" Then
            Data = .Details!POnum
            If Data = "None" Then
                .Details!POnum = .Op & "," & .txtPO.Text
                GoTo Next8
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!POnum = .Details!POnum & "," & .Op & "," & .txtPO.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!POnum
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt8
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt8:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!POnum = Data & .Op & "," & .txtPO.Text
                Else
                    .Details!POnum = Data & "," & .Op & "," & .txtPO.Text
                End If
            End If
        Else
            If .Details!POnum = "None" Or .Details!POnum = "" Then
                .Details!POnum = "None"
            End If
        End If
Next8:
        
        'Aircraft Number
        If .txtAircraft.Text <> "" Then
            Data = .Details!Aircraftnum
            If Data = "None" Then
                .Details!Aircraftnum = .Op & "," & .txtAircraft.Text
                GoTo Next9
            End If
            Open App.Path & "\Tem.txt" For Output As #1
                Print #1, Data
            Close #1
            Open App.Path & "\Tem.txt" For Input As #1
                Do Until EOF(1)
                    Input #1, strInput
                    Input #1, NI
                    If strInput = .Op Then
                        IsAValue = True
                        Exit Do
                    Else
                        IsAValue = False
                    End If
                Loop
            Close #1
            Kill App.Path & "\Tem.txt"
            If IsAValue = False Then
                .Details!Aircraftnum = .Details!Aircraftnum & "," & .Op & "," & .txtAircraft.Text
            Else
                Open App.Path & "\Text.txt" For Output As #1
                    Print #1, .Details!Aircraftnum
                Close #1
                Open App.Path & "\Text.txt" For Input As #1
                    Data = ""
                    Do Until EOF(1)
                        Input #1, It
                        If It = .Op Then GoTo FoundIt9
                        Data = Data & It & ","
                        Input #1, It
                        Data = Data & It & ","
                    Loop
FoundIt9:
                    Input #1, It
                    Do Until EOF(1)
                        Input #1, TheRest
                        Data = Data & TheRest & ","
                    Loop
                Close #1
                Kill App.Path & "\Text.txt"
                If Data <> "" Then Data = Left(Data, Len(Data) - 1)
                If Data = "" Then
                    .Details!Aircraftnum = Data & .Op & "," & .txtAircraft.Text
                Else
                    .Details!Aircraftnum = Data & "," & .Op & "," & .txtAircraft.Text
                End If
            End If
        Else
            If .Details!Aircraftnum = "None" Or .Details!Aircraftnum = "" Then
                .Details!Aircraftnum = "None"
            End If
        End If
Next9:
        
        .Details.Update
    End With
End Function

Public Function ClearDetails()
    With frmMain
        .txtAirbill.Text = ""
        .txtAirbill.BackColor = &H8000000F
        .txtAirbill.Enabled = False
        .txtAircraft.Text = ""
        .txtAircraft.BackColor = &H8000000F
        .txtAircraft.Enabled = False
        .txtDate.Text = ""
        .txtDate.BackColor = &H8000000F
        .txtDate.Enabled = False
        .txtDescription.Text = ""
        .txtDescription.BackColor = &H8000000F
        .txtDescription.Enabled = False
        .txtDestination.Text = ""
        .txtDestination.BackColor = &H8000000F
        .txtDestination.Enabled = False
        .txtOperations.Text = ""
        .txtOperations.BackColor = &H8000000F
        .txtOperations.Enabled = False
        .txtPO.Text = ""
        .txtPO.BackColor = &H8000000F
        .txtPO.Enabled = False
        .txtFlight.Text = ""
        .txtFlight.BackColor = &H8000000F
        .txtFlight.Enabled = False
        .txtWeight.Text = ""
        .txtWeight.BackColor = &H8000000F
        .txtWeight.Enabled = False
    End With
End Function

Public Function ClearFields()
    With frmMain
        .txtAirbill.Text = ""
        .txtAircraft.Text = ""
        .txtDate.Text = ""
        .txtDescription.Text = ""
        .txtDestination.Text = ""
        .txtOperations.Text = ""
        .txtPO.Text = ""
        .txtFlight.Text = ""
        .txtWeight.Text = ""
    End With
End Function

Public Function DBCreateTable(ByVal DBName1 As String, ByVal Table As String, Optional ByVal Password As String = ";pwd=")
     ' Add database table
     On Error GoTo AddError
     ' set database name to users passed name
     DBs = DBName1
     DBPassword = Password
     If Dir(DBs) = "" Then
        MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          If Left(Table, 1) = "[" And Right(Table, 1) = "]" Then
               DBTable = Table
          Else
               DBTable = "[" & Table & "]"
          End If
          ' open database
          Set Database1 = OpenDatabase(DBs, False, False, Password)
          ' create new table with one field called Key
          Database1.Execute "Create Table " & DBTable & " ([Key] number)"
          Database1.Close
     End If
     Exit Function
AddError:
     MsgBox "Error Deleting Table " & DBTable & "." & vbCrLf & "Or Table Does Not Exist.", vbExclamation, "Drop Error"
End Function
