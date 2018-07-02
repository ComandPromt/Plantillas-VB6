Attribute VB_Name = "mdlGetWinDirDeclare"
Public Declare Function GetWindowsDirectory Lib "kernel32" _
      Alias "GetWindowsDirectoryA" _
     (ByVal lpBuffer As String, ByVal nSize As Long) As Long ' Windows API function declaration
     
Public Function TrimNull(MyString As String)

' Input: The string representing the Windows directory
' Process: Even though this is not C or C++, that is what the API was written in.  A string in C/C++
'          always has a null terminator on the end of it to tell the C compiler where the end if the string is.
'          Visual Basic doesn't work this way, but the string passed in above still has a null terminator on it.
'          This function trims it away.
' Output: The newly de-nulled string

    Dim intPosition As Integer ' Integer representing the location of the null terminator in the string
   
    intPosition = InStr(MyString, Chr$(0)) ' Locating the null
    If intPosition Then ' Trimming it away
          TrimNull = Left(MyString, intPosition - 1)
    Else: TrimNull = MyString
    End If
  
End Function


