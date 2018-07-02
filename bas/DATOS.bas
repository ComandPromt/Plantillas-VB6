Attribute VB_Name = "DATOS"
Global temp As String
Sub COPYSTRU(DBFROM As Database, RGFROM As Recordset, DBNOMB As String)
    Dim tabla As String
    Dim c As String
    tabla = "CREATE TABLE " & App.Path & "\" & Trim(DBNOMB) & "("
    For i = 0 To RGFROM.Fields.Count - 1
        c = RGFROM.Fields(i).Name & " " & Mid(TipoCampo(RGFROM.Fields(i).Type), 3, Len(TipoCampo(RGFROM.Fields(i).Type))) & IIf(TipoCampo(RGFROM.Fields(i).Type) = "dbText" Or TipoCampo(RGFROM.Fields(i).Type) = "dbByte", " (" & RGFROM.Fields(i).Size & ")", "")
        If i <> RGFROM.Fields.Count - 1 Then
            tabla = tabla & c & ","
        Else
            tabla = tabla & c
        End If
    Next
    tabla = tabla + ");"
    DBFROM.Execute tabla
End Sub

Function TipoCampo(intTipo As Integer) As String

    Select Case intTipo
        Case dbBoolean
            TipoCampo = "dbBoolean"
        Case dbByte
            TipoCampo = "dbByte"
        Case dbInteger
            TipoCampo = "dbInteger"
        Case dbLong
            TipoCampo = "dbLong"
        Case dbCurrency
            TipoCampo = "dbCurrency"
        Case dbSingle
            TipoCampo = "dbSingle"
        Case dbDouble
            TipoCampo = "dbDouble"
        Case dbDate
            TipoCampo = "dbDateTime"
        Case dbText
            TipoCampo = "dbText"
        Case dbLongBinary
            TipoCampo = "dbLongBinary"
        Case dbMemo
            TipoCampo = "dbMemo"
        Case dbGUID
            TipoCampo = "dbGUID"
    End Select

End Function


Public Function NATEMP() As String
    Randomize
    NATEMP = "Q" + Trim(Str(Int((999999 * Rnd) + 1)))
End Function
