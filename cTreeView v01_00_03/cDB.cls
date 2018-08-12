VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Module Name:  mDB
' Author:       Graeme Grant
' Date:         27/09/2001
' Version:      01.00.01
' Description:  Private Database support routines
' Edit History: 00.01.00 14/10/2000 Initial Release
'               01.00.01 27/09/2001 Adjusted 'ExecuteSQL' to allow
'                                   ADODB.Command object to be configured
'                                   externally before execution.
'               01.00.01 27/09/2001 Add new NewID property.
'
'
'===========================================================================

Option Explicit

Public Enum eJetVersion
    ejvJet3 = 3
    ejvJet4 = 4
End Enum

Public Enum eFindRecord
    efrFindFirst = 1
    efrFindLast = 2
    efrFindNext = 3
    efrFindPrevious = 4
End Enum

Public Enum eMoveRecord
    emrMoveFirst = 1
    emrMoveLast = 2
    emrMoveNext = 3
    emrMovePrevious = 4
End Enum

Private moCon        As ADODB.Connection
Private msConnect    As String
Private mlPkID       As Long

Public Function Apostrophe(sFieldString As String) As String

    Dim lLen   As Long
    Dim lCount As Long
    Dim apostr As Long

    If InStr(sFieldString, "'") Then
        lLen = Len(sFieldString)
        lCount = 1

        Do While lCount <= lLen
            If Mid(sFieldString, lCount, 1) = "'" Then
                apostr = lCount
                sFieldString = Left(sFieldString, apostr) & "'" & _
                Right(sFieldString, lLen - apostr)
                lLen = Len(sFieldString)
                lCount = lCount + 1
            End If
            lCount = lCount + 1
        Loop

    End If
    Apostrophe = sFieldString

End Function

Public Function InitDB(ByVal FileName As String, _
              Optional ByVal User As String = "admin", _
              Optional ByVal Password As String = "", _
              Optional ByVal DefPath As String = "", _
              Optional ByVal JetVersion As eJetVersion = ejvJet3) As Boolean

    On Error GoTo ErrorHandler

    '-- Initialise module-level objects
    Set moCon = New ADODB.Connection

    If Len(Trim$(DefPath)) = 0 Then DefPath = App.Path + "\"
    Select Case JetVersion
        Case ejvJet3
            msConnect = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" + _
                         Trim$(FileName) + ";DefaultDir=" + Trim$(DefPath) + ";UID=" + _
                         Trim$(User) + ";PWD=;" + Trim$(Password)
        Case ejvJet4
            msConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "Data Source=" & Trim$(FileName) & ";" & _
                         "Jet OLEDB:Database Password=;" & _
                         "Jet OLEDB:Engine Type=5;"
    End Select

    moCon.Open msConnect
    InitDB = True
    Exit Function

ErrorHandler:
'    MsgBox Err.Source + " ERROR " + CStr(Err.Number) + " - " + Err.Description

End Function

Public Sub KillDB()
    If (Not moCon Is Nothing) Then Set moCon = Nothing
End Sub

Public Function CreateRS(oRs As ADODB.Recordset, SQL As String) As Boolean

    Dim oCmd As ADODB.Command

    Set oRs = New ADODB.Recordset
    Set oCmd = New ADODB.Command
    Set oCmd.ActiveConnection = moCon
    oCmd.CommandText = SQL
    With oRs
        .CursorLocation = ADODB.adUseClient
        .StayInSync = True
        .CacheSize = 1
        .Open oCmd, , ADODB.adOpenDynamic, ADODB.adLockBatchOptimistic
    End With
    CreateRS = True
    Set oCmd = Nothing

End Function

Public Function ExecuteSQL(Optional ByVal SSQL As String, _
                           Optional AdoCmd As ADODB.Command) As Boolean

    On Error GoTo ErrorHandler

    Dim oCmd  As ADODB.Command
    Set oCmd = New ADODB.Command
    Dim oRs As ADODB.Recordset

    ExecuteSQL = False

    If AdoCmd Is Nothing Then
        Set oCmd.ActiveConnection = moCon
        With oCmd
            .CommandType = adCmdText
            .CommandText = SSQL
        End With
        oCmd.Execute , , adExecuteNoRecords
    Else
        Set oCmd = AdoCmd
        Set oCmd.ActiveConnection = moCon
        oCmd.Execute , , adExecuteNoRecords
        Set oRs = moCon.Execute("SELECT @@Identity", , adCmdText)
        mlPkID = oRs(0).Value
    End If

    ExecuteSQL = True
    Set oCmd = Nothing

Exit Function

ErrorHandler:
'    gErrorHandler Err.Number, Err.Description, OBJNAME
End Function

Public Property Get NewID() As Long
    NewID = mlPkID
End Property

Public Function RecordCount(oRs As ADODB.Recordset) As Long

    Dim vBookmark As Variant

    On Error GoTo ErrorHandler

    With oRs
        vBookmark = .Bookmark
        .MoveFirst
        RecordCount = .RecordCount
        .Bookmark = vBookmark
    End With
    Exit Function

ErrorHandler:
    If Err.Number = 3021 Then RecordCount = 0
End Function

Public Function FindDB(Dir As eFindRecord, SSQL As String, oRs As ADODB.Recordset) As Boolean

    On Error GoTo ErrorHandler

    Dim vBookmark As Variant
    Dim sErrDesc  As String
    Dim lErrNo    As Long

    FindDB = False

    With oRs
        vBookmark = .Bookmark
        Select Case Dir
            Case efrFindFirst
                .MoveFirst
                .Find SSQL, , ADODB.adSearchForward '1
                If .EOF Then
                    .Bookmark = vBookmark
                    Exit Function
                End If

            Case efrFindLast
                .MoveLast
                .Find SSQL, , -1 'ADODB.adSearchBackward
                If .BOF Then
                    .Bookmark = vBookmark
                    Exit Function
                End If

            Case efrFindNext
                .Find SSQL, 1, 1 'ADODB.adSearchForward
                If .EOF Then
                    .Bookmark = vBookmark
                    Exit Function
                End If

            Case efrFindPrevious
                .Find SSQL, 1, -1 'ADODB.adSearchBackward
                If .BOF Then
                    .Bookmark = vBookmark
                    Exit Function
                End If

        End Select
    End With

'    moData.RS2Obj moRS
    FindDB = True
Exit Function

ErrorHandler:
'    If Err.Number <> 3021 Then
'        gErrorHandler Err.Number, Err.Description, OBJNAME
'    End If
End Function

Public Function MoveDB(Dir As eMoveRecord, oRs As ADODB.Recordset) As Boolean

    On Error GoTo ErrorHandler

    MoveDB = False

    With oRs
        Select Case Dir
            Case emrMoveFirst
                .MoveFirst
            Case emrMoveLast
                .MoveLast
            Case emrMoveNext
                .MoveNext
                If .EOF Then
                    .MoveLast
                    Exit Function
                End If
            Case emrMovePrevious
                .MovePrevious
                If .BOF Then
                    .MoveFirst
                    Exit Function
                End If
        End Select
        .Resync ADODB.adAffectCurrent ', .adResyncAllValues
    End With

'    moData.RS2Obj moRS
    MoveDB = True

Exit Function

ErrorHandler:
'    gErrorHandler Err.Number, Err.Description, OBJNAME

End Function
