Attribute VB_Name = "ReadWrite"
'ReadWrite.Bas
'An original program by LeGrev3@aol.com
'Submitted for downloading Dec 6, 2000
'A module that creates, writes to, and retrieves data from a .INI file
'Can be used to store and retrieve info like passwords and user preferences

Option Explicit

Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)

Public strMySystemFile As String
Public strSection As String
Public Const BUFF_SIZ As Long = 9160
Public Const READ_BUFF As Long = 255
'**********
Public strLoginName As String
Public strPassword As String
Public strColor As String
Public lngColor As Long
Public lngRetVal As Long


Sub Main()
    Dim strFileName As String
    strFileName = "MySystem.ini"        'set this to the file you want created
    
    strMySystemFile = App.Path & "\" & strFileName
    Form1.Show
End Sub

Function WriteToFile(strFileSection As String, strKey As String, strValue As String) As Long
'parameters: strFileSection - string used as a file subheader for every section
'            strKey - string used as key to write the value to file ( ex: UserName)
'            strValue - the string value to write to file ( ex: Password)
'returns:   -1 string to write is more than 255 chars
'            0 system write failure
'            1 write to file succesful

    If Len(strKey) > READ_BUFF Or Len(strValue) > READ_BUFF Then
        MsgBox "Can't write more than " & READ_BUFF & " characters for key or value."
        WriteToFile = -1
        Exit Function
    End If
    WriteToFile = WritePrivateProfileStringByKeyName(strFileSection, strKey, strValue, strMySystemFile)
End Function

Function ReadFromFile(strFileSection As String, strKey As String) As String
'parameters: strFileSection - string used as a file subheader for every section
'            strKey - string used as key to read the value from file ( ex: UserName)
'returns:    a string with the value tied to the key when written to file
'            a null string "" if Key is not on file

    Dim strValue As String
    Dim lngRetLen As Long
    
    strValue = String(READ_BUFF + 1, Space(1))
    lngRetLen = GetPrivateProfileStringByKeyName(strFileSection, strKey, "", strValue, READ_BUFF, strMySystemFile)
    If lngRetLen > 1 Then
        ReadFromFile = Left(strValue, lngRetLen)
    Else
        ReadFromFile = ""
    End If
End Function

Function DeleteFromFile(strFileSection As String, strKey As String) As Long
'parameters: strFileSection - string used as file subheader for every section
'            strKey - the string used in writing is also used for deleting
'returns:  -1 key or section is a null string
'           0 for system delete failure
'           1 for successfule delete

    If Len(strFileSection) = 0 Or Len(strKey) = 0 Then
        MsgBox "Null string parameter not allowed for DeleteFromFile."
        DeleteFromFile = -1
        Exit Function
    End If
    DeleteFromFile = WritePrivateProfileStringToDeleteKey(strFileSection, strKey, 0, strMySystemFile)
End Function

