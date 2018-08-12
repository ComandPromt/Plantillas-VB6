Attribute VB_Name = "modCode"
Option Explicit

Public dbNordwind As Database


Public rsKunden As Recordset
Public rsBestellung As Recordset
Public rsBestellDetails As Recordset


Public Function Datenbank() As Boolean
Dim dbPath As String
On Error GoTo dbErrors
dbPath = App.Path & "\Nordwind.mdb"
Set dbNordwind = DBEngine.Workspaces(0).OpenDatabase(dbPath, False)
Set rsBestellung = dbNordwind.OpenRecordset("Bestellungen", dbOpenTable)
Datenbank = True
Exit Function
dbErrors:
Datenbank = False
MsgBox (Err.Description)
End Function

