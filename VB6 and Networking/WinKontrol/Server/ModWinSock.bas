Attribute VB_Name = "ModWinSock"
Option Explicit
Public Inconnection As Boolean


Sub SendData(Data As String)
'Winsock SendData Command + Error Checking

    On Error GoTo ErrKontrol
    Dim TimeOut As Long
    FRMMain.Sock.SendData Data
    
    Do Until (FRMMain.Sock.State = 0) Or (TimeOut < 10000)
        DoEvents
        TimeOut = TimeOut + 1
        If TimeOut > 10000 Then Exit Do
    Loop
    
ErrKontrol:
Exit Sub
    
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'======================================================================
' (EvalData Function)
'
'  Purpose - Extract data from a given string, to the right or left
'            of a specified character.
'
'  Parameters:
'     sIncoming - The String you want to extract data from.
'     iRtLt     - Extract from the Left, 1.
'                 Extract from the right, 2.
'     sDivider  - The character that seperates the data in
'                 the string. <default = ",">
'
'  Returns:
'     (type)String
'     Returns the data to the right or left of strDivider.
'======================================================================
             
' THis function can be used to seperate endless bits of
' data sent withth SendData Function. Although it can get a little
' cumbersome with really long lists.

' If you would like more info on exactly how to accomplish this
' E-mail me or Message on ICQ and I will show you.

' THIS FUNCTION WAS WRITTEN BY gh0ul (gh0ul@hotmail.com)
' AND NOT BY ME, THE AUTHOR OF THIS PROGRAM

Public Function EvalData(sIncoming As String, iRtLt As Integer, _
                  Optional sDivider As String) As String
   Dim i As Integer
   Dim tempStr As String
   ' Storage for the current Divider
   Dim sSplit As String
   
   ' the current character used to divide the data
   If sDivider = "" Then
      sSplit = ","
   Else
      sSplit = sDivider
   End If
   
   ' getting the right or left?
   Select Case iRtLt
        
      Case 1
          ' remove the data to the Left of the Current Divider
          For i = 0 To Len(sIncoming)
            tempStr = Left(sIncoming, i)
            
            If Right(tempStr, 1) = sSplit Then
              EvalData = Left(tempStr, Len(tempStr) - 1)
              Exit Function
            End If
          Next
          
      Case 2
          ' remove the data to the Right of the Current Divider
          For i = 0 To Len(sIncoming)
            tempStr = Right(sIncoming, i)
            
            If Left(tempStr, 1) = sSplit Then
              EvalData = Right(tempStr, Len(tempStr) - 1)
              Exit Function
            End If
          Next
   End Select
   
End Function


