Attribute VB_Name = "iniCode"
Option Explicit

Public Function readINIFile(txtKeyname As String) As String
Dim lngRetVal As Long
Dim strTmp1 As String, strTmp2 As String, strDefault As String
Dim blnRetVal As Boolean

strTmp1 = App.Path & "\" & txtIniFile
strTmp2 = Space(200)
strDefault = ""

lngRetVal = GetPrivateProfileString("default", txtKeyname, strDefault, strTmp2, Len(strTmp2), strTmp1)

If lngRetVal = 0 Then
  Select Case txtKeyname
    Case "DatabaseLoc"
      blnRetVal = writeINIFile(txtKeyname, App.Path & "\" & "timesheets.mdb")
  End Select
  If blnRetVal = False Then
    'MsgBox "#GlobalCode #readINIFile Cannot write INI File first time, creating"
  Else
    Select Case txtKeyname
      Case "DatabaseLoc"
        readINIFile = App.Path & "\" & "timesheets.mdb"
    End Select
  End If
Else
  Debug.Print "Read Key " & txtKeyname & " Value " & Left(strTmp2, lngRetVal)
  readINIFile = Left(strTmp2, lngRetVal)
End If
End Function

Public Function writeINIFile(txtKeyname As String, txtKeyValue As String) As Boolean
Dim lngRetVal As Long
Dim strTmp As String

strTmp = App.Path & "\" & txtIniFile

lngRetVal = WritePrivateProfileString("default", txtKeyname, txtKeyValue, strTmp)

If lngRetVal = 0 Then
  MsgBox "Error Writing " & strTmp
  writeINIFile = False
Else
  Debug.Print "Wrote Key " & txtKeyname & " Value " & txtKeyValue & " INI File " & strTmp
  writeINIFile = True
End If
End Function


