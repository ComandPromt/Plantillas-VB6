VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'local variable(s) to hold property value(s)
Private mvarlngUserID As Long 'local copy
Private mvarstrFirstName As String 'local copy
Private mvarstrLastName As String 'local copy
Private mvarstrAbbreviation As String 'local copy
Private mvarbytSecurityLevel As Byte 'local copy

Public Sub deleteUser()
  Dim rstUser As New ADODB.Recordset
  
  cmdSelectUserByID.Parameters(0) = mvarlngUserID
  Set rstUser = returnRS(cmdSelectUserByID)
  If rstUser.EOF <> True Then
    rstUser.DELETE
    rstUser.UpdateBatch
  Else
    Debug.Print "#clsUser.deleteUser Error: cannot find user with lngUserID=" & mvarlngUserID
  End If
  rstUser.Close
  Set rstUser = Nothing
End Sub

Public Sub loadUser(lngUserID As Long)
  Dim rstUser As New ADODB.Recordset
  
  cmdSelectUserByID.Parameters(0) = lngUserID
  Set rstUser = returnRS(cmdSelectUserByID)
  
  If rstUser.EOF <> True Then
    mvarlngUserID = rstUser![lngUserID]
    mvarstrFirstName = rstUser![strFirstName]
    mvarstrLastName = rstUser![strLastName]
    mvarstrAbbreviation = rstUser![strAbbreviation]
    mvarbytSecurityLevel = rstUser![bytSecurityLevel]
  Else
    Debug.Print "#clsUser.loadUser Cannot find record for selected user ID"
  End If
  rstUser.Close
  Set rstUser = Nothing
End Sub

Public Sub clear()
    mvarlngUserID = 0
    mvarstrFirstName = ""
    mvarstrLastName = ""
    mvarstrAbbreviation = ""
    mvarbytSecurityLevel = 1
End Sub

Public Sub display()
  frmUserMaintenance.txtFirstName = mvarstrFirstName
  frmUserMaintenance.txtLastName = mvarstrLastName
  frmUserMaintenance.txtAbbreviation = mvarstrAbbreviation
  frmUserMaintenance.cboSecurityLevel.ListIndex = mvarbytSecurityLevel - 1
End Sub

Public Sub addUser()
  Dim rstUser As New ADODB.Recordset
  
  Set rstUser = returnRS(cmdSelectUsers)
  
  With rstUser
    .AddNew
      ![strFirstName] = mvarstrFirstName
      ![strLastName] = mvarstrLastName
      ![strAbbreviation] = mvarstrAbbreviation
      ![bytSecurityLevel] = mvarbytSecurityLevel
    .Update
  End With
  copyUser usrOld
  rstUser.Close
  Set rstUser = Nothing
End Sub

Public Sub editUser(lngUserID As Long)
  Dim rstUser As New ADODB.Recordset
  
  cmdSelectUserByID.Parameters(0) = lngUserID
  Set rstUser = returnRS(cmdSelectUserByID)
  If rstUser.EOF = False Then
    With rstUser
      ![strFirstName] = mvarstrFirstName
      ![strLastName] = mvarstrLastName
      ![strAbbreviation] = mvarstrAbbreviation
      ![bytSecurityLevel] = mvarbytSecurityLevel
      .Update
    End With
    copyUser usrOld
  Else
    Debug.Print "#clsUser.editUser Cannot find record for selected user ID"
  End If
  rstUser.Close
  Set rstUser = Nothing
End Sub

Public Sub copyUser(usrTmp As clsUser)
    usrTmp.lngUserID = mvarlngUserID
    usrTmp.strFirstName = mvarstrFirstName
    usrTmp.strLastName = mvarstrLastName
    usrTmp.strAbbreviation = mvarstrAbbreviation
    usrTmp.bytSecurityLevel = mvarbytSecurityLevel
End Sub

Public Property Let bytSecurityLevel(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngUserID = 5
    mvarbytSecurityLevel = vData
End Property


Public Property Get bytSecurityLevel() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngUserID
    bytSecurityLevel = mvarbytSecurityLevel
End Property


Public Property Let lngUserID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngUserID = 5
    mvarlngUserID = vData
End Property


Public Property Get lngUserID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngUserID
    lngUserID = mvarlngUserID
End Property

Public Property Let strFirstName(ByVal vData As String)
    mvarstrFirstName = vData
End Property


Public Property Get strFirstName() As String
    strFirstName = mvarstrFirstName
End Property

Public Property Let strLastName(ByVal vData As String)
    mvarstrLastName = vData
End Property


Public Property Get strLastName() As String
    strLastName = mvarstrLastName
End Property

Public Property Let strAbbreviation(ByVal vData As String)
    mvarstrAbbreviation = vData
End Property


Public Property Get strAbbreviation() As String
    strAbbreviation = mvarstrAbbreviation
End Property

