Attribute VB_Name = "modMisc"
'-------------------------------------------------------------------------------
' Copyright © 2001 Gregory Kirk. All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute the
' Application Files (and/or any modified version) in any way you find useful,
' provided that you agree that Gregory Kirk has no warranty, obligations or
' liability for any Application Files.
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' This module contains misc functions required for working with ini file types,
' adding an entry into the log, and file exist function.
'-------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------------
'APIs for reading and writing retrieving information from an ini file:
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'-------------------------------------------------------------------------------
'Function for reading a key from an ini file:
Public Function ReadINI(iSection As String, iKey As String, iniFile As String)
Dim RetStr As String, Retlen As String, iPath As String
iPath = App.Path & "\" & iniFile
RetStr = Space$(255)
Retlen = GetPrivateProfileString(iSection, iKey, "", RetStr, Len(RetStr), iPath)
RetStr = Left$(RetStr, Retlen)
ReadINI = IIf(RetStr = "", "Empty...", RetStr)
End Function

'-------------------------------------------------------------------------------
'Function for writing a key to an ini file:
Public Sub WriteINI(iSection As String, iKey As String, iniFile As String, Text As String)
WritePrivateProfileString iSection, iKey, Text, iniFile
End Sub

'-------------------------------------------------------------------------------
'Function for determining if a specified file exists:
Public Function FileExist(Filename As String) As Boolean
  On Error Resume Next
  FileExist = (Dir$(Filename) <> "")
End Function

'-------------------------------------------------------------------------------
'Function for adding an entry to the log:
Public Sub rLog(Addr As String, LocP As String, RemP As String, Status As String, TimeStamp As String, Optional ByPass As Boolean)
Dim Item As ListItem, i As Integer

'ByPass is for when you do not want an entry to go through the validation loop
'to determine if the entry all ready exists.
'Currently the entry for blocking uses "ByPass". I suppose I could creat a seperate
'log for blocking but for now...
If ByPass = True Then GoTo ByPassLoop

'//Checks to see if the entry is already there.
For i = 1 To frmMain.lvwLog.ListItems.Count
    If Addr = frmMain.lvwLog.ListItems(i).Text And LocP = frmMain.lvwLog.ListItems(i).ListSubItems(1) And Status = frmMain.lvwLog.ListItems(i).ListSubItems(3) Then
    Exit Sub
    End If
Next
ByPassLoop:

Set Item = frmMain.lvwLog.ListItems.Add()
Item.Text = Addr 'Host Address (IP or Name alias)
Item.SubItems(1) = LocP 'Local Port
Item.SubItems(2) = RemP 'Remote Port
Item.SubItems(3) = Status 'Connection Status
Item.SubItems(4) = TimeStamp
Form1.Text1.Text = Addr
Form1.Text2.Text = Status
Form1.Text3.Text = TimeStamp

Form1.Show
    'MsgBox Addr & "  " & Status & "  " & TimeStamp

End Sub

'//Pause for the specified duration. Duration is in seconds.
'Public Sub Pause(duration)
'Dim Current As Long
'Current = Timer
'Do Until Timer - Current >= duration
'    DoEvents
'Loop
'End Sub

