Attribute VB_Name = "securityCode"
Option Explicit

Public Sub removeLoggedInUser(lngUserID As Long)
  Dim rstUser As New ADODB.Recordset

  cmdSelectLoggedInUserByID.Parameters(0) = lngUserID
  Set rstUser = returnRS(cmdSelectLoggedInUserByID)
  If rstUser.EOF <> True Then
    rstUser.DELETE
    rstUser.UpdateBatch
  Else
    MsgBox "#securityCode#removeLoggedInUser Error: Cannot find user ID " & lngUserID
  End If
  rstUser.Close
  Set rstUser = Nothing
End Sub

Public Sub addLoggedInUser(lngUserID As Long)
  Dim rstUsers As New ADODB.Recordset

  Set rstUsers = returnRS(cmdSelectLoggedInUsers)
  With rstUsers
    .AddNew
      ![lngUserID] = lngUserID
    .Update
  End With
  rstUsers.Close
  Set rstUsers = Nothing
End Sub

Public Function userLoggedIn(lngUserID As Long) As Boolean
  Dim rstUser As New ADODB.Recordset

  cmdSelectLoggedInUserByID.Parameters(0) = lngUserID
  Set rstUser = returnRS(cmdSelectLoggedInUserByID)
  If rstUser.EOF <> True Then
    userLoggedIn = True
  Else
    userLoggedIn = False
  End If
  rstUser.Close
  Set rstUser = Nothing
End Function

Public Function usersLoggedIn(lngUserID As Long) As Boolean
  Dim rstUsers As New ADODB.Recordset

  Set rstUsers = returnRS(cmdSelectLoggedInUsers)
  If rstUsers.EOF <> True Then
    rstUsers.MoveFirst
    While rstUsers.EOF <> True
      If rstUsers![lngUserID] <> lngUserID Then
        usersLoggedIn = True
        rstUsers.MoveLast
      End If
      rstUsers.MoveNext
    Wend
  Else
    usersLoggedIn = False
  End If
  rstUsers.Close
  Set rstUsers = Nothing
End Function
