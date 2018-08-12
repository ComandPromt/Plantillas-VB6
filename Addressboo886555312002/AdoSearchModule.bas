Attribute VB_Name = "AdoSearchModule"
'********************************************************************
'~Created by Adam Lankford 4/18/2001
'********************************************************************
Option Explicit

Public Function Search(parameter As String, rs As ADODB.Recordset, X As Field) As Boolean
    Dim foundFlag As Boolean
    Dim i As Integer

    With rs
        If .RecordCount > 0 Then
            .MoveFirst
                For i = 1 To .RecordCount
                    If X = parameter Then
                        foundFlag = True
                        i = .RecordCount
                    End If
                    If foundFlag = False Then
                        .MoveNext
                    End If
                Next i
                If foundFlag = True Then
                   'MsgBox ("Record has been location!")
                Else
                    MsgBox ("No Contact found!")
                    foundFlag = False
                    .MoveFirst
                End If
        Else
            MsgBox ("There are no records To search!")
            foundFlag = False
        End If
    End With
    Search = foundFlag
End Function




