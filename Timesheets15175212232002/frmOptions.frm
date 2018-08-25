VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      ToolTipText     =   "Click to exit"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      ToolTipText     =   "Click to accept changes"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtDatabaseLocation 
      Height          =   330
      HelpContextID   =   1110
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Location of the database containing the process maps"
      Top             =   120
      Width           =   3030
   End
   Begin VB.CommandButton btnBrowseDatabase 
      Caption         =   "Browse..."
      Height          =   330
      HelpContextID   =   1110
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Backend Database Location"
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   165
      Width           =   2310
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strOldDatabaseLoc As String

Private Sub btnBrowseDatabase_Click()
  frmSaveAs.CommonDialog1.Filter = "Database Files (*.mdb) | *.mdb|All Files (*.*) | *.*"
  frmSaveAs.CommonDialog1.InitDir = parseDirectory(Me.txtDatabaseLocation)
  frmSaveAs.CommonDialog1.ShowOpen
  If Len(frmSaveAs.CommonDialog1.FileName) > 0 Then
    Me.txtDatabaseLocation = frmSaveAs.CommonDialog1.FileName
  End If
End Sub

Private Sub btnExit_Click()
  Unload Me
End Sub

Private Sub btnOK_Click()
  Dim blnSuccess As Boolean
  Dim strTmpConnection As String
  Dim cnnTmp As New ADODB.Connection
  Dim rstTmp As New ADODB.Recordset
  Dim intIndex As Integer, intSelectedTab As Integer
  Dim strMapName As String
  blnSuccess = True
  '
  'Check Database Location
  '
  If Len(Me.txtDatabaseLocation) < 5 Then
    MsgBox "Database Location Invalid, Please Reselect", vbOKOnly, "File Location Error"
    blnSuccess = False
  Else
    If Right(Me.txtDatabaseLocation, 4) <> ".mdb" Then
      MsgBox "You have not selected a valid database file, Please Reselect", vbOKOnly, "File Type Error"
      blnSuccess = False
    End If
  End If
  '
  'Establish a connection to the database and check the table names to check if it appears to have the correct
  'structure
  '
  If blnSuccess <> False And strOldDatabaseLoc <> Me.txtDatabaseLocation Then
    strTmpConnection = strConnectionStart & Me.txtDatabaseLocation & strConnectionEnd
    cnnTmp.ConnectionString = strTmpConnection
    cnnTmp.CursorLocation = adUseClient
    cnnTmp.Open strTmpConnection, "Admin"
    
    On Error GoTo errorHandler
    '
    'Query each table and see if it exists, if so, the database is probably valid
    '
    Set rstTmp = cnnTmp.Execute("[Project Details]", , adCmdTable)
    Set rstTmp = cnnTmp.Execute("[Kilometers]", , adCmdTable)
    Set rstTmp = cnnTmp.Execute("[Times]", , adCmdTable)
    Set rstTmp = cnnTmp.Execute("[User Details]", , adCmdTable)
    Set rstTmp = Nothing
    blnSuccess = writeINIFile("DatabaseLoc", Me.txtDatabaseLocation)
    If blnSuccess <> True Then
      MsgBox "ERROR: timesheets.ini file could not be written to, database location not updated", vbOKOnly, "Error Writing INI File"
      Exit Sub
    End If
    '
    'Successful, close the connection, redefine the string, and reopen it
    '
    strConnection = strConnectionStart & Me.txtDatabaseLocation & strConnectionEnd
    DETimesheets.Close
    DETimesheets.ConnectionString = strConnection
    DETimesheets.CursorLocation = adUseClient
    DETimesheets.Open strConnection, "Admin"
    globalCode.defineCommandObjects
  Else
    If blnSuccess = True Then
    End If
  End If

  Exit Sub
errorHandler:
  MsgBox "ERROR: The database you selected does not appear to be valid, please reselect", vbCritical, "Database Error"
  Set rstTmp = Nothing
  Exit Sub

End Sub

Private Sub Form_Load()
    Me.txtDatabaseLocation = readINIFile("DatabaseLoc")
    strOldDatabaseLoc = Me.txtDatabaseLocation
End Sub
