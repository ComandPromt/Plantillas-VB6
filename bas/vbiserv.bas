Attribute VB_Name = "modVIServices"
' Service GUIDs
Public Const ControlContextGuid = "{D18C18C1-304C-11D0-8158-00A0C91BBEE3}"
Public Const AppObjectGuid = "{0C539790-12E4-11CF-B661-00AA004CD6D8}"
Public Const UrlBuilderGuid = "{73CEF3D9-AE85-11CF-A406-00AA00C00940}"
Public Const VIServiceGuid = "{5492AFA0-6D81-11d0-B746-0000F81E081D}"

' Returns a service object from the host
Public Declare Function GetService Lib "vbiserv.dll" (vControl As Variant, ByVal strGuid As String, oService As Variant) As Boolean

' ODBC API declares
Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hDbc As Long, pHstmt As Long) As Integer
Declare Function SQLPrepare Lib "odbc32.dll" Alias "SQLExecDirect" (ByVal hStmt As Long, ByVal szSqlStr As String, ByVal cbSqlString As Integer) As Integer
Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hStmt As Long, ByVal fOption As Integer) As Integer
Declare Function SQLNumResultCols Lib "odbc32.dll" (ByVal hStmt As Long, pccol As Integer) As Integer
Declare Function SQLColAttribute Lib "odbc32.dll" (ByVal hStmt As Long, ByVal icol As Integer, ByVal iID As Integer, ByVal szBuffer As String, ByVal cbBuffer As Integer, pcbBuffer As Integer, pNumericAttr As Long) As Integer

' ODBC Return codes
Public Const conSqlInvalidHandle = -2
Public Const conSqlError = -1
Public Const conSqlSuccess = 0
Public Const conSqlSuccessWithInfo = 1
Public Const conSqlStillExecuting = 2
Public Const conSqlNeedData = 99
Public Const conSqlNoDataFound = 100

' Miscellaneous constants
Public Const conSQLNTS = -3

' ODBC constants for SQLFreeStmt
Public Const conSqlClose = 0
Public Const conSqlDrop = 1
Public Const conSqlUnbind = 2
Public Const conSqlResetParams = 3

' Constants for SQLColAttribute
Public Const SqlDescName = 1011
Public Const SqlDescLabel = 18
Public Const SqlDescType = 1002
Public Const SqlDescTypeName = 14
Public Const SqlDescLength = 1003

