VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClientAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'// CA Error Code
    Private Const CWB_OK                As Long = 0
    Private Const CLASS_NAME            As String = "CClientAccess"

    Private mvarDatabases               As Collection
    Private mvarProvider                As String
    Private mvarDefaultPackageLibrary   As String
    Private mvarUserID                  As String
    Private mvarPassword                As String
    Private mvarServer                  As String
    Private mvarSystem                  As String

Public Property Set Databases(ByVal vData As Collection)
Attribute Databases.VB_Description = "Collection of databases to use when creating the ADODB.Connection object."

    Set mvarDatabases = vData
End Property

Public Property Get Databases() As Collection

    Set Databases = mvarDatabases
End Property

Public Function Connect() As Object
Attribute Connect.VB_Description = "Returns an ADODB.Connection reference "
Attribute Connect.VB_UserMemId = 0

    On Error GoTo Connect_Error

    Dim conNew                          As Object
    Dim clsCALogon                      As CCALogOn
    
'// Ensure valid property assignment before attempting connection
    If ValidateProperties Then
    
'// Create the new object
        Set conNew = CreateObject("ADODB.Connection")
    
'// Use the CA Logon Class to cache NT log-on information
        Set clsCALogon = New CCALogOn
        If clsCALogon.LogonToMachine(mvarServer, mvarUserID, mvarPassword) <> CWB_OK Then
            Err.Raise CA_LOGON_FAIL, CLASS_NAME, "Could not log on to server"
        End If
    
'// We have cached user information, attempt to open the connection
        conNew.CommandTimeout = 0
        conNew.ConnectionTimeout = 0
        conNew.Open BuildConnectionString

'// Return the object
        Set Connect = conNew
    
'// Destroy CA object
        Set clsCALogon = Nothing

    End If

    Exit Function
    
Connect_Error:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Property Let System(ByVal vData As String)
Attribute System.VB_Description = "The system to connect to when creating the ADODB.Connection object"

    mvarSystem = vData
End Property

Public Property Get System() As String

    System = mvarSystem
End Property

Public Property Let Server(ByVal vData As String)
Attribute Server.VB_Description = "The server to connect to when creating the ADODB.Connection object"

    mvarServer = vData
End Property

Public Property Get Server() As String

    Server = mvarServer
End Property

Public Property Let Password(ByVal vData As String)
Attribute Password.VB_Description = "The password to use when creating the ADODB.Connection object"

    mvarPassword = vData
End Property

Public Property Get Password() As String

    Password = mvarPassword
End Property

Public Property Let UserID(ByVal vData As String)
Attribute UserID.VB_Description = "The User ID to use when creating the ADODB.Connection object. SAAUSER is the default"

    mvarUserID = vData
End Property

Public Property Get UserID() As String

    UserID = mvarUserID
End Property

Public Property Let DefaultPackageLibrary(ByVal vData As String)
Attribute DefaultPackageLibrary.VB_Description = "The default package library of the AS/400 system. QGPL is the default value"

    mvarDefaultPackageLibrary = vData
End Property

Public Property Get DefaultPackageLibrary() As String

    DefaultPackageLibrary = mvarDefaultPackageLibrary
End Property

Public Property Let Provider(ByVal vData As String)
Attribute Provider.VB_Description = "The provider to use when creating the ADODB.Connection object"

    mvarProvider = vData
End Property

Public Property Get Provider() As String

    Provider = mvarProvider
End Property

Private Sub Class_Initialize()

'// Create new collection
    Set mvarDatabases = New Collection
    
'// Set some default properties
    mvarDefaultPackageLibrary = "QGPL"
    mvarUserID = ""
    mvarPassword = ""

End Sub

Private Sub Class_Terminate()

    Set mvarDatabases = Nothing
    
End Sub

Public Function BuildConnectionString() As String

    Dim strX            As String
    Dim vItem           As Variant
    
'// Build general string
    strX = "ODBC;"
    strX = strX & "Provider=" & mvarProvider & ";"
    strX = strX & "Driver={Client Access ODBC Driver (32-bit)};"
    strX = strX & "DFTPKGLIB=" & mvarDefaultPackageLibrary & ";"
    strX = strX & "Server=" & mvarServer & ";"
    strX = strX & "System=" & mvarSystem & ";"
    strX = strX & "uid=" & mvarUserID & ";"
    
'// Password if appropriate
    If Len(mvarPassword) > 0 Then
        strX = strX & "pwd=" & mvarPassword & ";"
    End If
    
'// Databases to connect to, truncate string to remove extra commas
    strX = strX & "Database="
    For Each vItem In mvarDatabases
        strX = strX & vItem & ", "
    Next
    strX = Left(strX, Len(strX) - 2)
    
'// Return value
    BuildConnectionString = strX

End Function

Private Function ValidateProperties() As Boolean

    If Len(mvarProvider) = 0 Then
        Err.Raise CA_NO_PROVIDER, CLASS_NAME, "Specify a provider for the connection"
    End If
    
    If Len(mvarServer) = 0 Then
        Err.Raise CA_NO_SERVER, CLASS_NAME, "Specify a server for the connection"
    End If
    
    If Len(mvarSystem) = 0 Then
        Err.Raise CA_NO_SYSTEM, CLASS_NAME, "Specify a system for the connection"
    End If
    
    If (mvarDatabases.Count = 0) Then
        Err.Raise CA_NO_DATABASE, CLASS_NAME, "You must specify at least one database for the connection"
    End If
    
    ValidateProperties = True
    
End Function
