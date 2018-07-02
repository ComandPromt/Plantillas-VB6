Attribute VB_Name = "modMunir"
Option Explicit

Global Const DBName As String = "biblio.mdb"
Global strPath As String
Global db As Database
Global rsTitles As Recordset
Global rsAuteur As Recordset

Public Sub GeefDetails(source As String)
    Dim strSQL As String
    
    Screen.MousePointer = vbHourglass
    strSQL = "SELECT Titles.Title, Titles.ISBN, Authors.Author"
    strSQL = strSQL & " FROM Titles INNER JOIN (Authors INNER JOIN [Title Author] ON Authors.Au_ID = [Title Author].Au_ID) ON Titles.ISBN = [Title Author].ISBN"
    strSQL = strSQL & " WHERE (((Titles.Title) = '" & source & "'))"
    Set rsAuteur = db.OpenRecordset(strSQL)
    If Not (rsAuteur.BOF And rsAuteur.EOF) Then
        frmMunir2.Caption = source
        rsAuteur.MoveLast
        If rsAuteur.RecordCount > 1 Then Call ShowCommand(True) Else Call ShowCommand(False)
        rsAuteur.MoveFirst
        Call InvullenGegevens
    Else
        frmMunir2.Caption = "niets gevonden"
        frmMunir2.Label1(0).Caption = ""
        frmMunir2.Label1(1).Caption = ""
    End If
    Screen.MousePointer = vbNormal
    
End Sub
Public Sub InvullenGegevens()
    frmMunir2.Label1(0).Caption = rsAuteur.Fields(1).Value
    frmMunir2.Label1(1).Caption = rsAuteur.Fields(2).Value
    
End Sub
Public Sub ShowCommand(zichtbaar As Boolean)
    Dim bX As Byte
    
    If zichtbaar Then
        For bX = 0 To 3
            frmMunir2.Command1(bX).Visible = True
        Next bX
    Else
        For bX = 0 To 3
            frmMunir2.Command1(bX).Visible = False
        Next bX
    End If
    
End Sub
