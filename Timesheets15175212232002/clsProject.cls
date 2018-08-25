VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarlngProjectID As Long 'local copy
Private mvarstrProjectNumber As String 'local copy
Private mvarstrProjectName As String 'local copy
Private mvarstrProjectDescription As String 'local copy
Private mvarstrCustomer As String
Private mvardatCreated As Date 'local copy
Private mvardatClosed As Date 'local copy
Private mvarlngCreatedByID As Long 'local copy
Private mvarlngManagerID As Long 'local copy
Private mvarstrManagerName As String
Private mvarstrCreatedByName As String
Private mvarcurBudgetLabour As Currency 'local copy
Private mvarcurBudgetMaterial As Currency 'local copy
Private mvarcurBudgetTravel As Currency 'local copy

Public Sub deleteProject()
  Dim rstProject As New ADODB.Recordset
  
  cmdSelectProjectByID.Parameters(0) = mvarlngProjectID
  Set rstProject = returnRS(cmdSelectProjectByID)
  If rstProject.EOF <> True Then
    rstProject.DELETE
    rstProject.UpdateBatch
  Else
    Debug.Print "#clsProject.deleteProject Error: cannot find project with lngProjectID=" & mvarlngProjectID
  End If
  rstProject.Close
  Set rstProject = Nothing
End Sub

Public Sub loadProject(lngProjectID As Long)
  Dim rstProject As New ADODB.Recordset
  
  cmdSelectProjectByID.Parameters(0) = lngProjectID
  Set rstProject = returnRS(cmdSelectProjectByID)
  
  If rstProject.EOF <> True Then
    mvarlngProjectID = rstProject![lngProjectID]
    mvarstrProjectNumber = rstProject![strProjectNumber]
    mvarstrProjectName = rstProject![strProjectName]
    mvarstrProjectDescription = rstProject![memProjectDescription]
    mvarstrCustomer = rstProject![strCustomer]
    mvardatCreated = rstProject![datCreated]
    mvardatClosed = rstProject![datClosed]
    mvarlngCreatedByID = rstProject![lngCreatedByID]
    mvarlngManagerID = rstProject![lngManagerID]
    mvarcurBudgetLabour = rstProject![curBudgetLabour]
    mvarcurBudgetMaterial = rstProject![curBudgetMaterial]
    mvarcurBudgetTravel = rstProject![curBudgetTravel]
  Else
    Debug.Print "#clsProject.loadProject Cannot find record for selected project ID"
  End If
  rstProject.Close
  Set rstProject = Nothing
End Sub
 
Public Function getLatestProject() As Long
  Dim rstProjects As New ADODB.Recordset
  
  Set rstProjects = returnRS(cmdSelectProjectsOrderByID)
  
  If rstProjects.EOF <> True Then
    rstProjects.MoveLast
    getLatestProject = rstProjects![lngProjectID]
  Else
    Debug.Print "#clsProject.getLatestProject Cannot find any project records"
  End If
  rstProjects.Close
  Set rstProjects = Nothing
End Function

Public Sub clear()
  mvarlngProjectID = 0
  mvarstrProjectNumber = ""
  mvarstrProjectName = ""
  mvarstrProjectDescription = ""
  mvarstrCustomer = ""
  mvardatCreated = 0
  mvardatClosed = 0
  mvarlngCreatedByID = 0
  mvarlngManagerID = 0
  mvarcurBudgetLabour = 0
  mvarcurBudgetMaterial = 0
  mvarcurBudgetTravel = 0
End Sub

Public Sub display()
  frmProjectMaintenance.txtProjectNumber = mvarstrProjectNumber
  frmProjectMaintenance.txtProjectName = mvarstrProjectName
  frmProjectMaintenance.txtProjectDescription = mvarstrProjectDescription
  frmProjectMaintenance.txtCustomer = mvarstrCustomer
  If mvardatCreated = 0 Then
    frmProjectMaintenance.txtDateCreated = ""
  Else
    frmProjectMaintenance.txtDateCreated = mvardatCreated
  End If
  If mvardatClosed = 0 Then
    frmProjectMaintenance.txtDateClosed = ""
  Else
    frmProjectMaintenance.txtDateClosed = mvardatClosed
  End If
  If mvarlngCreatedByID > 0 Then
    frmProjectMaintenance.cboCreatedBy.ListIndex = projectMaintenanceCode.matchUserID(frmProjectMaintenance.cboCreatedBy, mvarlngCreatedByID)
  Else
    frmProjectMaintenance.cboCreatedBy.ListIndex = -1
  End If
  If mvarlngManagerID > 0 Then
    frmProjectMaintenance.cboManager.ListIndex = projectMaintenanceCode.matchUserID(frmProjectMaintenance.cboManager, mvarlngManagerID)
  Else
    frmProjectMaintenance.cboManager.ListIndex = -1
  End If
  frmProjectMaintenance.txtEstimatedLabour = mvarcurBudgetLabour
  frmProjectMaintenance.txtEstimatedMaterial = mvarcurBudgetMaterial
  frmProjectMaintenance.txtEstimatedTravel = mvarcurBudgetTravel
End Sub

Public Sub addProject()
  Dim rstProject As New ADODB.Recordset
  
  Set rstProject = returnRS(cmdSelectProjects)
  
  With rstProject
    .AddNew
    ![strProjectNumber] = mvarstrProjectNumber
    ![strProjectName] = mvarstrProjectName
    ![memProjectDescription] = mvarstrProjectDescription
    ![strCustomer] = mvarstrCustomer
    ![datCreated] = mvardatCreated
    ![datClosed] = mvardatClosed
    ![lngCreatedByID] = mvarlngCreatedByID
    ![lngManagerID] = mvarlngManagerID
    ![curBudgetLabour] = mvarcurBudgetLabour
    ![curBudgetMaterial] = mvarcurBudgetMaterial
    ![curBudgetTravel] = mvarcurBudgetTravel
    .Update
  End With
  copyProject proOld
  rstProject.Close
  Set rstProject = Nothing
End Sub

Public Sub editProject(lngProjectID As Long)
  Dim rstProject As New ADODB.Recordset
  
  cmdSelectProjectByID.Parameters(0) = lngProjectID
  Set rstProject = returnRS(cmdSelectProjectByID)
  If rstProject.EOF = False Then
    With rstProject
      ![strProjectNumber] = mvarstrProjectNumber
      ![strProjectName] = mvarstrProjectName
      ![memProjectDescription] = mvarstrProjectDescription
      ![strCustomer] = mvarstrCustomer
      ![datCreated] = mvardatCreated
      ![datClosed] = mvardatClosed
      ![lngCreatedByID] = mvarlngCreatedByID
      ![lngManagerID] = mvarlngManagerID
      ![curBudgetLabour] = mvarcurBudgetLabour
      ![curBudgetMaterial] = mvarcurBudgetMaterial
      ![curBudgetTravel] = mvarcurBudgetTravel
      .Update
    End With
    copyProject proOld
  Else
    Debug.Print "#clsProject.editProject Cannot find record for selected project ID"
  End If
  rstProject.Close
  Set rstProject = Nothing
End Sub

Public Sub copyProject(proTmp As clsProject)
    proTmp.lngProjectID = mvarlngProjectID
    proTmp.strProjectNumber = mvarstrProjectNumber
    proTmp.strProjectName = mvarstrProjectName
    proTmp.strProjectDescription = mvarstrProjectDescription
    proTmp.strCustomer = mvarstrCustomer
    proTmp.datCreated = mvardatCreated
    proTmp.datClosed = mvardatClosed
    proTmp.lngCreatedByID = mvarlngCreatedByID
    proTmp.lngManagerID = mvarlngManagerID
    proTmp.curBudgetLabour = mvarcurBudgetLabour
    proTmp.curBudgetMaterial = mvarcurBudgetMaterial
    proTmp.curBudgetTravel = mvarcurBudgetTravel
End Sub

Public Property Let curBudgetTravel(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.curBudgetTravel = 5
    mvarcurBudgetTravel = vData
End Property


Public Property Get curBudgetTravel() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.curBudgetTravel
    curBudgetTravel = mvarcurBudgetTravel
End Property



Public Property Let curBudgetMaterial(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.curBudgetMaterial = 5
    mvarcurBudgetMaterial = vData
End Property


Public Property Get curBudgetMaterial() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.curBudgetMaterial
    curBudgetMaterial = mvarcurBudgetMaterial
End Property



Public Property Let curBudgetLabour(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.curBudgetLabour = 5
    mvarcurBudgetLabour = vData
End Property


Public Property Get curBudgetLabour() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.curBudgetLabour
    curBudgetLabour = mvarcurBudgetLabour
End Property





Public Property Let lngManagerID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngManagerID = 5
    mvarlngManagerID = vData
End Property


Public Property Get lngManagerID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngManagerID
    lngManagerID = mvarlngManagerID
End Property

Public Property Let strManagerName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCreatedByID = 5
    mvarstrManagerNamee = vData
End Property


Public Property Get strManagerName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCreatedByID
    strManagerName = mvarstrManagerName
End Property

Public Property Let strCustomer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCreatedByID = 5
    mvarstrCustomer = vData
End Property


Public Property Get strCustomer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCreatedByID
    strCustomer = mvarstrCustomer
End Property

Public Property Let lngCreatedByID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCreatedByID = 5
    mvarlngCreatedByID = vData
End Property


Public Property Get lngCreatedByID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCreatedByID
    lngCreatedByID = mvarlngCreatedByID
End Property

Public Property Let strCreatedByName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCreatedByID = 5
    mvarstrCreatedByName = vData
End Property


Public Property Get strCreatedByName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCreatedByID
    strCreatedByName = mvarstrCreatedByName
End Property

Public Property Let datClosed(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.datClosed = 5
    mvardatClosed = vData
End Property


Public Property Get datClosed() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.datClosed
    datClosed = mvardatClosed
End Property



Public Property Let datCreated(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.datCreated = 5
    mvardatCreated = vData
End Property


Public Property Get datCreated() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.datCreated
    datCreated = mvardatCreated
End Property



Public Property Let strProjectDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strProjectDescription = 5
    mvarstrProjectDescription = vData
End Property


Public Property Get strProjectDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strProjectDescription
    strProjectDescription = mvarstrProjectDescription
End Property



Public Property Let strProjectName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strProjectName = 5
    mvarstrProjectName = vData
End Property


Public Property Get strProjectName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strProjectName
    strProjectName = mvarstrProjectName
End Property



Public Property Let strProjectNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strProjectNumber = 5
    mvarstrProjectNumber = vData
End Property


Public Property Get strProjectNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strProjectNumber
    strProjectNumber = mvarstrProjectNumber
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




