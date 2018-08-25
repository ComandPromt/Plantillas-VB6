VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'local variable(s) to hold property value(s)
Private mvarlngID As Long 'local copy
Private mvarlngUserID As Long 'local copy
Private mvarlngProjectID As Long 'local copy
Private mvardatDate As Date 'local copy
Private mvardblHours As Double 'local copy
Private mvarstrDescription As String 'local copy
Private mvarblnInvoiced As Boolean 'local copy

Public Sub loadTime(lngID As Long)
  Dim rstTime As New ADODB.Recordset
  
  cmdSelectTimeByID.Parameters(0) = lngID
  Set rstTime = returnRS(cmdSelectTimeByID)
  
  If rstTime.EOF <> True Then
    mvarlngID = rstTime![lngID]
    mvarlngUserID = rstTime![lngUserID]
    mvarlngProjectID = rstTime![lngProjectID]
    mvarstrDescription = rstTime![memDescription]
    mvardatDate = rstTime![datDate]
    mvardblHours = rstTime![dblHours]
    mvarblnInvoiced = rstTime![blnInvoiced]
  Else
    Debug.Print "#clsTime.loadTime Cannot find record for selected time ID"
  End If
  rstTime.Close
  Set rstTime = Nothing
End Sub

Public Sub deleteTime(lngID As Long)
  Dim rstTime As New ADODB.Recordset
  
  cmdSelectTimeByID.Parameters(0) = lngID
  Set rstTime = returnRS(cmdSelectTimeByID)
  
  If rstTime.EOF <> True Then
    rstTime.DELETE
    rstTime.Update
  Else
    Debug.Print "#clsTime.deleteTime Cannot find record for selected time ID"
  End If
  rstTime.Close
  Set rstTime = Nothing
End Sub

Public Sub clear()
    mvarlngID = 0
    mvarlngUserID = usrLoggedIn.lngUserID
    mvarlngProjectID = 0
    mvarstrDescription = ""
    mvardatDate = 0
    mvardblHours = 0
    mvarblnInvoiced = False
End Sub

Public Sub display()
  timesheetCode.selectCorrectProject mvarlngProjectID, frmTimeSheet.cboProject
  frmTimeSheet.txtDescription = mvarstrDescription
  frmTimeSheet.txtTime = mvardblHours
End Sub

Public Sub addTime(lngProjectID As Long)
  Dim rstTime As New ADODB.Recordset
  
  cmdSelectUserProjectTimes.Parameters(0) = mvarlngProjectID
  cmdSelectUserProjectTimes.Parameters(1) = mvarlngUserID
  Set rstTime = returnRS(cmdSelectUserProjectTimes)
  
  With rstTime
    .AddNew
       ![lngUserID] = mvarlngUserID
       ![lngProjectID] = mvarlngProjectID
       ![memDescription] = mvarstrDescription
       ![datDate] = datSelectedDate
       ![dblHours] = mvardblHours
       ![blnInvoiced] = False
    .Update
  End With
  copyTime timOld
  rstTime.Close
  Set rstTime = Nothing
End Sub

Public Sub editTime(lngID As Long)
  Dim rstTime As New ADODB.Recordset
  
  cmdSelectTimeByID.Parameters(0) = lngID
  Set rstTime = returnRS(cmdSelectTimeByID)
  If rstTime.EOF = False Then
    With rstTime
       ![lngProjectID] = mvarlngProjectID
       ![datDate] = mvardatDate
       ![memDescription] = mvarstrDescription
       ![dblHours] = mvardblHours
       .Update
    End With
    copyTime timOld
  Else
    Debug.Print "#clsTime.editTime Cannot find record for selected time ID"
  End If
  rstTime.Close
  Set rstTime = Nothing
End Sub

Public Sub copyTime(timTmp As clsTime)
  timTmp.lngID = mvarlngID
  timTmp.lngUserID = mvarlngUserID
  timTmp.lngProjectID = mvarlngProjectID
  timTmp.strDescription = mvarstrDescription
  timTmp.datDate = mvardatDate
  timTmp.dblHours = mvardblHours
  timTmp.blnInvoiced = mvarblnInvoiced
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

Public Property Let dblHours(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngUserID = 5
    mvardblHours = vData
End Property

Public Property Get dblHours() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngUserID
    dblHours = mvardblHours
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
