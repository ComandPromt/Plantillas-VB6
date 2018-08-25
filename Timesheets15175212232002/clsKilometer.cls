VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKilometer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'local variable(s) to hold property value(s)
Private mvarlngID As Long 'local copy
Private mvarlngUserID As Long 'local copy
Private mvarlngProjectID As Long 'local copy
Private mvardatDate As Date 'local copy
Private mvardblKilometers As Double 'local copy
Private mvarstrDescription As String 'local copy
Private mvarblnInvoiced As Boolean 'local copy
Private mvarblnReimbursed As Boolean 'local copy

Public Sub loadKilometer(lngID As Long)
  Dim rstKilometer As New ADODB.Recordset
  
  cmdSelectKilometerByID.Parameters(0) = lngID
  Set rstKilometer = returnRS(cmdSelectKilometerByID)
  
  If rstKilometer.EOF <> True Then
    mvarlngID = rstKilometer![lngID]
    mvarlngUserID = rstKilometer![lngUserID]
    mvarlngProjectID = rstKilometer![lngProjectID]
    mvarstrDescription = rstKilometer![memDescription]
    mvardatDate = rstKilometer![datDate]
    mvardblKilometers = rstKilometer![dblKilometers]
    mvarblnInvoiced = rstKilometer![blnInvoiced]
    mvarblnReimbursed = rstKilometer![blnReimbursed]
  Else
    Debug.Print "#clsKilometer.loadKilometer Cannot find record for selected kilometer ID"
  End If
  rstKilometer.Close
  Set rstKilometer = Nothing
End Sub

Public Sub deleteKilometer(lngID As Long)
  Dim rstKilometer As New ADODB.Recordset
  
  cmdSelectKilometerByID.Parameters(0) = lngID
  Set rstKilometer = returnRS(cmdSelectKilometerByID)
  
  If rstKilometer.EOF <> True Then
    rstKilometer.DELETE
    rstKilometer.Update
  Else
    Debug.Print "clsKilometer.deleteKilometer Cannot find record for selected kilometer ID"
  End If
  rstKilometer.Close
  Set rstKilometer = Nothing
End Sub

Public Sub clear()
    mvarlngID = 0
    mvarlngUserID = usrLoggedIn.lngUserID
    mvarlngProjectID = 0
    mvarstrDescription = ""
    mvardatDate = 0
    mvardblKilometers = 0
    mvarblnInvoiced = False
    mvarblnReimbursed = False
End Sub

Public Sub display()
  frmKilometers.txtDate = mvardatDate
  frmKilometers.txtKilometers = mvardblKilometers
  frmKilometers.txtDescription = mvarstrDescription
End Sub

Public Sub addKilometer(lngProjectID As Long)
  Dim rstKilometer As New ADODB.Recordset
  
  cmdSelectUserProjectKilometers.Parameters(0) = lngProjectID
  cmdSelectUserProjectKilometers.Parameters(1) = mvarlngUserID
  Set rstKilometer = returnRS(cmdSelectUserProjectKilometers)
  
  With rstKilometer
    .AddNew
       ![lngUserID] = mvarlngUserID
       ![lngProjectID] = lngProjectID
       ![memDescription] = mvarstrDescription
       ![datDate] = mvardatDate
       ![dblKilometers] = mvardblKilometers
       ![blnInvoiced] = False
       ![blnReimbursed] = False
    .Update
  End With
  copyKilometer klmOld
  rstKilometer.Close
  Set rstKilometer = Nothing
End Sub

Public Sub editKilometer(lngID As Long)
  Dim rstKilometer As New ADODB.Recordset
  
  cmdSelectKilometerByID.Parameters(0) = lngID
  Set rstKilometer = returnRS(cmdSelectKilometerByID)
  If rstKilometer.EOF = False Then
    With rstKilometer
       ![datDate] = mvardatDate
       ![memDescription] = mvarstrDescription
       ![dblKilometers] = mvardblKilometers
       .Update
    End With
    copyKilometer klmOld
  Else
    Debug.Print "#clsKilometer.editKilometer Cannot find record for selected time ID"
  End If
  rstKilometer.Close
  Set rstKilometer = Nothing
End Sub

Public Sub copyKilometer(klmTmp As clsKilometer)
  klmTmp.lngID = mvarlngID
  klmTmp.lngUserID = mvarlngUserID
  klmTmp.lngProjectID = mvarlngProjectID
  klmTmp.strDescription = mvarstrDescription
  klmTmp.datDate = mvardatDate
  klmTmp.dblKilometers = mvardblKilometers
  klmTmp.blnInvoiced = mvarblnInvoiced
  klmTmp.blnReimbursed = mvarblnReimbursed
End Sub

Public Property Let lngID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngUserID = 5
    mvarlngID = vData
End Property

Public Property Get lngID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngUserID
    lngID = mvarlngID
End Property

Public Property Let lngProjectID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngProjectID = 5
    mvarlngProjectID = vData
End Property


Public Property Get lngProjectID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngProjectID
    lngProjectID = mvarlngProjectID
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

Public Property Let strDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strProjectDescription = 5
    mvarstrDescription = vData
End Property


Public Property Get strDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strProjectDescription
    strDescription = mvarstrDescription
End Property

Public Property Let datDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.datCreated = 5
    mvardatDate = vData
End Property


Public Property Get datDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.datCreated
    datDate = mvardatDate
End Property

Public Property Let dblKilometers(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngUserID = 5
    mvardblKilometers = vData
End Property

Public Property Get dblKilometers() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngUserID
    dblKilometers = mvardblKilometers
End Property

Public Property Let blnInvoiced(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnExternal = 5
    mvarblnInvoiced = vData
End Property


Public Property Get blnInvoiced() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnExternal
    blnInvoiced = mvarblnInvoiced
End Property

Public Property Let blnReimbursed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnExternal = 5
    mvarblnReimbursed = vData
End Property

Public Property Get blnReimbursed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnExternal
    blnReimbursed = mvarblnReimbursed
End Property
