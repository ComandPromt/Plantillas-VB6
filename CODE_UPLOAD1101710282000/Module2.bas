Attribute VB_Name = "Module1"
Public db As Database
Public rs4 As Recordset
Public X As Long
Public Y As Long
Public Grosseur As Long
Public Function OpenConnection()
If oConnection.State <> adStateOpen Then
    oConnection.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=YourODBCconnection"
End If
End Function
Public Sub Initialise()
Set db = OpenDatabase("lelecteur\labasededonnée")
Set rs4 = db.OpenRecordset(LaRequete)
End Sub
Public Function Eval_Valeur(leField As String, LaRequete As String) As Variant

'strsql = "select " & leField & " from " & LaRequete & ";"
If Not rs4.EOF Then
    Eval_Valeur = rs4(leField)
Else
    Eval_Valeur = ""
End If

End Function
