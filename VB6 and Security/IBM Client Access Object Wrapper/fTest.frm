VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "TEST CONNECT"
      Height          =   795
      Left            =   960
      TabIndex        =   0
      Top             =   1020
      Width           =   2595
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()

    Dim conNew                  As Object
    Dim rs                      As Object
    Dim clsClientAccess         As ClientAccess.CClientAccess

'// Create connection object
    Set clsClientAccess = New ClientAccess.CClientAccess
    
'// Connection properties
    With clsClientAccess
        .DefaultPackageLibrary = "QGPL"
        .Provider = "IBMDA400"
        .UserID = "YOURUSERID"
        .Password = "YOURPASSWORD"
        .Server = "SERVERNAME"
        .System = "SYSTEMNAME"
        .Databases.Add "LIBRARYNAME1"
        Set conNew = .Connect
    End With
    
'// Create recordset object
    Set rs = CreateObject("ADODB.Recordset")
    
'// Open recordset
    rs.ActiveConnection = conNew
    rs.Locktype = 1
    rs.CursorType = 0
    rs.Source = "select * from libraryname.filename"
    rs.Open
    
'// Destroy objects
    rs.Close
    Set rs = Nothing
    
    conNew.Close
    Set conNew = Nothing

End Sub

