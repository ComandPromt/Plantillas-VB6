Attribute VB_Name = "Module1"

Public IsDelOrBack As Boolean
Public Function AutoComplete(TheText As TextBox, TheDB As DataBase, TheTable As String, TheField As String) As Boolean
On Error Resume Next
Dim OldLen As Integer
Dim dsTemp As Recordset
AutoComplete = False
If Not TheText.Text = "" And IsDelOrBack = False Then
OldLen = Len(TheText.Text)
    Set dsTemp = TheDB.OpenRecordset("Select * from " & TheTable & " where " & TheField & " like '" & TheText.Text & "*'", dbOpenDynaset)
      If Err = 3075 Then
      End If
         If Not dsTemp.RecordCount = 0 Then
            TheText.Text = dsTemp(TheField)
                If TheText.SelText = "" Then
                    TheText.SelStart = OldLen
                Else
                    TheText.SelStart = InStr(TheText.Text, TheText.SelText)
                End If
                    TheText.SelLength = Len(TheText.Text)
                    AutoComplete = True
        End If

End If
End Function

Public Function CheckIsDelOrBack(TheKey As Integer) As Boolean
    If TheKey = vbKeyBack Or TheKey = vbKeyDelete Then
        IsDelOrBack = True
        CheckIsDelOrBack = True
    Else
        IsDelOrBack = False
        CheckIsDelOrBack = False
    End If
End Function


