Attribute VB_Name = "basNet"
Option Explicit

Private Declare Function WNetGetUser Lib "mpr" Alias _
   "WNetGetUserA" (ByVal lpName As String, _
   ByVal lpUserName As String, lpnLength As Long) As Long

'
'  Returns the user name or "" if the
'  user is not logged on.
'
Public Function NetUserName() As String
   Dim i As Long
   Dim UserName As String * 255

   i = WNetGetUser("", UserName, 255)
   
   If i = 0 Then
      NetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
   Else
      NetUserName = ""
   End If
   
End Function
